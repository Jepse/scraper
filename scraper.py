import os
import requests
from bs4 import BeautifulSoup
import openpyxl
import re
import validators
from urllib.parse import urljoin
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import asksaveasfilename
import threading
import sys
import traceback

def log_uncaught_exceptions(ex_cls, ex, tb):
    with open('app_errors.log', 'a') as f:
        f.write(''.join(traceback.format_tb(tb)))
        f.write('{0}: {1}\n'.format(ex_cls, ex))

sys.excepthook = log_uncaught_exceptions

def scrape_and_save(start_urls, output_excel_file, max_pages=100):
    visited_urls = set()
    urls_to_visit = set(start_urls)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["URL", "Email"])

    while urls_to_visit and len(visited_urls) < max_pages:
        url = urls_to_visit.pop()

        if url not in visited_urls:
            print("Scraping:", url)
            try:
                response = requests.get(url, timeout=10)
                soup = BeautifulSoup(response.text, "html.parser")

                # Find all links
                links = [a['href'] for a in soup.find_all('a', href=True)]
                for link in links:
                    full_url = urljoin(url, link)

                    # This condition ensures that we don't add URLs beyond our max limit
                    if len(visited_urls) + len(urls_to_visit) >= max_pages:
                        break

                    if validators.url(full_url) and full_url not in visited_urls:
                        urls_to_visit.add(full_url)

                # Find all email addresses
                emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', response.text)
                for email in emails:
                    ws.append([url, email])

            except Exception as e:
                print(f"Failed to scrape {url}: {e}")

            finally:
                visited_urls.add(url)

    wb.save(output_excel_file)
    print("Data saved to", output_excel_file)



def on_scrape():
    # Disable the button to prevent multiple threads running simultaneously
    scrape_button.config(state=tk.DISABLED)

    # Start the progress bar
    progress.start(10)

    # Use threading to avoid blocking the GUI
    threading.Thread(target=execute_scrape).start()

def execute_scrape():
    # Get the URL from the input field
    url = url_input.get()
    
    # Check if URL is valid
    if not validators.url(url):
        messagebox.showerror("Error", "Invalid URL.")
        reset_ui()
        return
    
    # Get output excel file location
    output_file = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    
    # If a file location was provided
    if output_file:
        try:
            # Call your scrape_and_save function
            scrape_and_save([url], output_file, 100)
            messagebox.showinfo("Success", f"Data saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to scrape: {e}")
    reset_ui()

def reset_ui():
    # Stop the progress bar and re-enable the scrape button
    progress.stop()
    scrape_button.config(state=tk.NORMAL)

# Create main window
root = tk.Tk()
root.title("Email Scraper")
root.geometry("600x300")

# Correctly determine the path for the .ico file when bundled
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle/exe
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

icon_path = os.path.join(base_path, 'Employai.ico')
root.iconbitmap(icon_path)

# Create label, input and button
url_label = ttk.Label(root, text="URL Input")
url_label.pack(pady=40)

url_input = ttk.Entry(root, width=50)
url_input.pack(pady=10)

scrape_button = ttk.Button(root, text="Scrape It", command=on_scrape)
scrape_button.pack(pady=10)

# Indeterminate progress bar
progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=300, mode='indeterminate')
progress.pack(pady=20)

root.mainloop()














