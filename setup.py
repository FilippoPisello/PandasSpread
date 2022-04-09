from setuptools import find_packages, setup

with open("readme.md", "r") as file:
    long_description = file.read()

setup(
    name="spreadpandas",
    version="0.0.1",
    description="Get the spreadsheet representation of a pandas data frame.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/FilippoPisello/SpreadPandas",
    author="Filippo Pisello",
    author_email="filippo.pisello@live.it",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Operating System :: OS Independent",
    ],
    packages=find_packages(),
    include_package_data=True,
    install_requires=["numpy>=1.10", "pandas>=1.1.0"],
    python_requires=">=3.7",
)
