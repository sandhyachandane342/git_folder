# git_folder
def generate_markdown_links(file_names, base_url):
    markdown_content = "### Uploaded Files\n\n"
    for file_name in file_names:
        markdown_content += f"- [{file_name}]({base_url}/{file_name})\n"
        print(markdown_content)
file_names = [
    "easyecom_suborders.py",
    "order_details_wh_holidays_date.py",
    "order_shipping_page_count.py",
    "read_new_file.py",
    "zec.py",
    "zmto_new_data_file.py",
    "zre.py"
]

base_url = "https://github.com/sandhyachandane342/git_folder"

markdown_text = generate_markdown_links(file_names, base_url)

with open("README.md", "w", encoding="utf-8") as markdown_file:
    markdown_file.write(markdown_text)

print("Markdown file created successfully.")
