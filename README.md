# git_folder
def generate_markdown_links(files):
    markdown = ""
    for file in files:
        file_url = f"https://github.com/sandhyachandane342/git_folder/blob/main/{file}"
        markdown += f"- [{file}]({file_url})\n"
    return markdown

uploaded_files = ["easyecom_suborders.py", 
                  "order_details_wh_holidays_date.py", 
                  "order_shipping_page_count.py", 
                  "read_new_file.py", 
                  "zec.py", 
                  "zmto_new_data_file.py", 
                  "zre.py"]

markdown_content = generate_markdown_links(uploaded_files)

print(markdown_content)
