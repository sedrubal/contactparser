#!/usr/bin/env python

"""
The setup script for the contactparser
"""

from setuptools import setup, find_packages
import os
import contactparser


def read(fname):
    """
    returns the content of the relative file fname
    """
    return open(os.path.join(os.path.dirname(__file__), fname)).read()

setup(
    name="contactparser",
    packages=find_packages(),
    entry_points={
        "console_scripts": ["contactparser = contactparser.contactparser:main"]
    },
    author=contactparser.__authors__,
    author_email=contactparser.__email__,
    license=contactparser.__license__,
    description=contactparser.__doc__,
    long_description=read("README.md"),
    url=contactparser.__url__,
    version=contactparser.__version__,
    install_requires=["BeautifulSoup4", "argparse", "argcomplete"]
)
