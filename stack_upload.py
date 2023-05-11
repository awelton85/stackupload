import fitz
import concurrent.futures
import re
import env
import time
from tkinter import filedialog
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright

SEARCH_WORDS = {"limestone", "granite", "stone", "kasota", "coldspring"}
# SEARCH_WORDS = {"limestone", "granite", "stone", "kasota", "brick", "st1", "st-1", "st2", "st-2", "coldspring"}


def search_and_highlight_page(args: tuple) -> None:
    page, words = args
    print(
        f"\nSearching page {page.number} of {len(fitz_document)}, "
        f"{page.number / len(fitz_document) * 100:.2f}% complete")
    for word in words:
        text_instances = page.search_for(word)
        for inst in text_instances:
            page.add_highlight_annot(inst)


# get input PDF path from user
def get_input_pdf_path() -> str:
    return filedialog.askopenfilename(
        initialdir="~/Downloads",
        title="Select A File",
        filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))


# create a fitz.Document object
def create_fitz_document(fitz_filepath: str) -> fitz.Document:
    try:
        return fitz.Document(fitz_filepath)
    except TypeError:
        print("No file selected")
        exit()


# save marked PDF to output PDF path
def save_marked_pdf(document: fitz.Document, marked_pdf_filename: str) -> None:
    print("\nSaving PDF..")
    document.save(marked_pdf_filename, garbage=4, deflate=True, clean=True)
    document.close()
    print("Finished saving marked PDF")


# open stack, log in, create new project, and upload marked PDF to new project
def upload_to_stackct(output_path: str) -> None:
    print("Starting upload to StackCT...")
    job_name = output_path.split("/")[-2]  # sets job name to the name of the folder containing the PDF

    def run(pw: Playwright) -> None:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.goto("https://www.stackct.com/")
        page.get_by_role("button", name="Sign In").first.click()
        page.get_by_role("paragraph").filter(has_text=re.compile(r"^Takeoff & Estimating$")).\
            get_by_role("link", name="Takeoff & Estimating").click()
        page.get_by_label("Business Email").click()
        page.get_by_label("Business Email").fill(env.EMAIL)
        page.get_by_role("button", name="Continue").click()
        page.get_by_label("Password").click()
        page.get_by_label("Password").fill(env.PASSWORD)
        page.get_by_role("button", name="Login").click()
        page.get_by_role("button", name="New Project").click()
        page.get_by_label("Project Name:").click()
        page.get_by_label("Project Name:").fill(job_name)
        page.get_by_role("button", name="Create and Launch").click()
        page.get_by_role("button", name="Local Files").click()
        with page.expect_file_chooser() as fc:
            page.get_by_role("button", name="Choose a local file").click()
            file_chooser = fc.value
            file_chooser.set_files(output_path)
            page.get_by_role("button", name="Done").click()
        print("Waiting for upload to complete...")
        time.sleep(20)

    with sync_playwright() as playwright:
        run(playwright)

    print("Done!\n")


# main
if __name__ == "__main__":
    original_file_path = get_input_pdf_path()  # get input PDF path from user
    start_time = datetime.now()  # start performance timer
    marked_filepath = original_file_path[:-4] + "_marked.pdf"  # create output PDF filepath
    fitz_document = create_fitz_document(original_file_path)  # create fitz.Document object from input PDF

    # searches and highlights all pages of PDF for words in SEARCH_WORDS using multithreading
    with concurrent.futures.ThreadPoolExecutor() as executor:
        executor.map(
            search_and_highlight_page, [(page, SEARCH_WORDS) for page in fitz_document]
        )

    save_marked_pdf(fitz_document, marked_filepath)  # save marked PDF to output PDF path
    end_time1 = datetime.now()  # end performance timer for writing PDF
    print(f"Search/write duration: {end_time1 - start_time}\n")

    upload_to_stackct(marked_filepath)  # uploads marked PDF to StackCT as a new project
    end_time2 = datetime.now()  # end performance timer for uploading PDF

    print(f"Upload duration: {end_time2 - end_time1}")
    print(f"Total duration: {end_time2 - start_time}")
