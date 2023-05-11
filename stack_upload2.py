import fitz
import re
import env
import time
import threading
from tkinter import filedialog
from datetime import datetime
from playwright.sync_api import Playwright, sync_playwright

SEARCH_WORDS = {"limestone", "granite", "stone", "kasota", "coldspring"}


# SEARCH_WORDS = {"limestone", "granite", "stone", "kasota", "brick", "st1", "st-1", "st2", "st-2", "coldspring"}


def search_and_highlight_page(start, finish, wordlist: set, document: fitz.Document) -> None:
    start.wait()
    for page in document:
        print(f"\nSearching page {page.number} of {len(fitz_document)}, "
              f"{page.number / len(fitz_document) * 100:.2f}% complete")
        for word in wordlist:
            text_instances = page.search_for(word)
            for inst in text_instances:
                page.add_highlight_annot(inst)
    save_marked_pdf(fitz_document, marked_filepath)
    finish.set()


# get input PDF path from user
def get_input_pdf_path() -> str:
    return filedialog.askopenfilename(
        initialdir="~/Downloads",
        title="Select A File",
        filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")),
    )


# create a fitz.Document object
def create_fitz_document(fitz_filepath: str) -> fitz.Document:
    try:
        return fitz.Document(fitz_filepath)
    except TypeError:
        print("No file selected")
        exit()


# save marked PDF to output PDF path
def save_marked_pdf(document: fitz.Document, marked_pdf_filename: str) -> None:
    print("\nSaving PDF...")
    document.save(marked_pdf_filename, garbage=4, deflate=True, clean=True)
    document.close()
    print("Finished saving marked PDF")


# open stack, log in, create new project, and upload marked PDF to new project
def upload_to_stackct(start, finish, output_path: str) -> None:
    start.set()
    job_name = output_path.split("/")[-2]  # sets job name to the name of the folder containing the PDF

    def run(pw: Playwright) -> None:
        browser = pw.chromium.launch(headless=False)
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
        page.get_by_label("Project Name:").fill(job_name)
        page.get_by_role("button", name="Create and Launch").click()
        page.get_by_role("button", name="Local Files").click()
        print("\nStarting upload...")
        finish.wait()
        with page.expect_file_chooser() as fc:
            page.get_by_role("button", name="Choose a local file").click()
            file_chooser = fc.value
            file_chooser.set_files(output_path)
            page.get_by_role("button", name="Done").click()
        print("Waiting for upload to complete...")
        time.sleep(5)

    with sync_playwright() as playwright:
        run(playwright)

    print("Done!\n")


# main
if __name__ == "__main__":
    original_file_path = get_input_pdf_path()  # get input PDF path from user
    start_time = datetime.now()  # start performance timer
    marked_filepath = (original_file_path[:-4] + "_marked.pdf")  # create output PDF filepath
    fitz_document = create_fitz_document(original_file_path)  # create fitz.Document object from input PDF

    # search and highlight all pages of PDF for words in SEARCH_WORDS using multithreading
    start_event = threading.Event()
    finish_event = threading.Event()
    thread2 = threading.Thread(target=upload_to_stackct,
                               args=(start_event, finish_event, marked_filepath))
    thread1 = threading.Thread(target=search_and_highlight_page,
                               args=(start_event, finish_event, SEARCH_WORDS, fitz_document))

    thread2.start()
    thread1.start()

    thread1.join()
    thread2.join()

    print(f"\nTotal duration: {datetime.now() - start_time}")
