# git_folder
import markdown2
markdown2.markdown("*boo!*")  # or use `html = markdown_path(PATH)`
'<p><em>boo!</em></p>\n'

 from markdown2 import Markdown
 markdowner = Markdown()
 markdowner.convert("*boo!*")
'<p><em>boo!</em></p>\n'
 markdowner.convert("**boom!**")
'<p><strong>boom!</strong></p>\n'
