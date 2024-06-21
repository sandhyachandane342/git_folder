# git_folder
from setuptools import setup
from pathlib import Path
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text()
setup(
    name='an_example_package',
    long_description=long_description,
    long_description_content_type='text/markdown'
)
