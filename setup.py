from setuptools import setup, find_packages
from os import path

here = path.abspath(path.dirname(__file__))

with open(here + "/README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="ontodev-axle",
    version="0.0.1",
    description="AXLE for XLSX",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ontodev/axle",
    author="Rebecca C Jackson",
    author_email="rbca.jackson@gmail.com",
    license="",
    classifiers=[
        "Programming Language :: Python :: 3",
        "Environment :: Console",
        "Operating System :: OS Independent",
        "License :: OSI Approved :: BSD License",
    ],
    packages=find_packages(exclude="tests"),
    python_requires=">=3.6, <4",
    entry_points={"console_scripts": ["axle=axle.cli:main"]},
)
