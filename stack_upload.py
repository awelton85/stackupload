import fitz
import concurrent.futures
import re
import env
import time
from tkinter import filedialog
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright, expect

# words to search for, edit this set as needed
SEARCH_WORDS = {"limestone", "granite", "stone", "kasota", "coldspring"}
# SEARCH_WORDS = {"limestone", "granite", "stone", "kasota", "brick", "st1", "st-1", "st2", "st-2", "coldspring"}


# find search words in document and highlight them
def search_and_highlight_page(args):
    page, words = args
    print(
        f"\nSearching page {page.number} of {len(doc)}, {page.number / len(doc) * 100:.2f}% complete"
    )
    for word in words:
        text_instances = page.search_for(word)
        for inst in text_instances:
            page.add_highlight_annot(inst)


# get path to pdf file
filename = filedialog.askopenfilename(
    initialdir="~/Downloads",
    title="Select A File",
    filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")),
)


# Create a Document object
try:
    doc = fitz.Document(filename)
except TypeError:
    print("No file selected")
    exit()

# Search and highlight
start_time = datetime.now()
with concurrent.futures.ThreadPoolExecutor() as executor:
    executor.map(search_and_highlight_page, [(page, SEARCH_WORDS) for page in doc])

# Save the document
marked_filename = filename[:-4] + "_marked.pdf"
print("\nWriting...")
doc.save(marked_filename, garbage=4, deflate=True, clean=True)
doc.close()
end_time1 = datetime.now()

print("\nFinished writing marked PDF")
print(f"Write duration: {end_time1 - start_time}\n")
print("Starting upload to StackCT...")

JOBNAME = marked_filename.split("/")[-2]


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.stackct.com/")
    page.get_by_role("button", name="Sign In").first.click()
    page.get_by_role("paragraph").filter(
        has_text=re.compile(r"^Takeoff & Estimating$")
    ).get_by_role("link", name="Takeoff & Estimating").click()
    page.get_by_label("Business Email").click()
    page.get_by_label("Business Email").fill(env.EMAIL)
    page.get_by_role("button", name="Continue").click()
    page.get_by_label("Password").click()
    page.get_by_label("Password").fill(env.PASSWORD)
    page.get_by_role("button", name="Login").click()
    page.get_by_role("button", name="New Project").click()
    page.get_by_label("Project Name:").click()
    page.get_by_label("Project Name:").fill(JOBNAME)
    page.get_by_role("button", name="Create and Launch").click()
    page.get_by_role("button", name="Local Files").click()
    with page.expect_file_chooser() as fc:
        page.get_by_role("button", name="Choose a local file").click()
        file_chooser = fc.value
        file_chooser.set_files(marked_filename)
        page.get_by_role("button", name="Done").click()
    print("Waiting for upload to complete...")
    time.sleep(10)

    # ---------------------
    context.close()
    browser.close()
    print("Done!\n")


with sync_playwright() as playwright:
    run(playwright)

end_time2 = datetime.now()
print(f"Upload duration: {end_time2 - end_time1}")
print(f"Total duration: {end_time2 - start_time}")
