# git_folder
def generate_markdown_file_list(files):
    markdown_content = "### Uploaded Files\n\n"
    for file_name in files:
        markdown_content += f"- [{file_name}]({file_name})\n"
    return markdown_content

uploaded_files = [
    "easyecom_suborders.py", 
    "order_details_wh_holidays_date.py", 
    "order_shipping_page_count.py", 
    "read_new_file.py", 
    "zec.py", 
    "zmto_new_data_file.py", 
    "zre.py"
]

markdown_text = generate_markdown_file_list(uploaded_files)
with open("uploaded_files.md", "w", encoding="utf-8") as markdown_file:
    markdown_file.write(markdown_text)

print("Markdown file created successfully.")
