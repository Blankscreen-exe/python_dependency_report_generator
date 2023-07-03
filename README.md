# Python Dependency Report Generator

It produces a `docx` format document which lists out all the dependencies that a certain python environments uses. For proper usage please the usage guides

## Usage

1. Make sure that your target virtual environment is already created and accessed, otherwise the report will be generated for the global environment.
2. place the `gen_docx.py` file in (preferably) your root directory of your project (or it could be any other directory for that matter)
3. install the dependency for this file which is just `python-docx`.
```shell
pip3 install python-docx
```
> DO NOT CONFUSE `python-docx` with `docx`. `docx` is an old implementation of `python-docx` and will cause problems for recent versions of python.

4. For basic usage, execute the `gen_docx.py` file using
```shell
python3 ./gen_docx.py 
```
5. and you will see a file named `dependencies_report.docx` created in the same directory. 
6. For some customizations you can use
```python
# ./your_file.py

from gen_docs import generate_dependency_docx as gdd

# initialize object
document = gdd()

# set text contents
document.set_title("My Title")
document.set_description("My short description.")

# get package information
dependencies = document.get_package_info()

# generate report
document.generate_report(dependencies)
```

## Preview
The generated document will look something like this
![Alt text](docs/img/preview.png)

## Further Development
I'm working on making this more customizable and user friendly. Might even make a little framework out of it. Feel free to let me know what you would like to see in this project.

- [ ] add setter methods for theme customizations
- [ ] add layout customizations
- [ ] convert it into a python package
