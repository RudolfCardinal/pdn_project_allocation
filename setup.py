#!/usr/bin/env/python

"""
setup.py

===============================================================================

    Copyright (C) 2019 Rudolf Cardinal (rudolf@pobox.com).

    This file is part of pdn_project_allocation.

    This is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This software is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this software. If not, see <https://www.gnu.org/licenses/>.

===============================================================================

Python package configuration.

"""

from setuptools import setup, find_packages

from pdn_project_allocation.version import VERSION

setup(
    name="pdn_project_allocation",
    version=VERSION,
    description="Allocate students to projects",
    url="https://github.com/RudolfCardinal/pdn_project_allocation.git",
    author="Rudolf Cardinal",
    author_email="rudolf@pobox.com",
    license="GNU General Public License v3 or later (GPLv3+)",
    # See https://pypi.org/classifiers/
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Education",
        "License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)",  # noqa
        "Natural Language :: English",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Topic :: Education",
    ],
    # Python code:
    packages=find_packages(),
    # Static files:
    # https://stackoverflow.com/questions/11848030/how-include-static-files-to-setuptools-python-package
    package_data={
        "pdn_project_allocation.testdata": ["*"],
    },
    # Requirements:
    install_requires=[
        "cardinal_pythonlib==1.1.23",
        "mip==1.14.1",
        "matching==1.4",
        "openpyxl==3.0.10",
        "lxml==4.9.1",  # Will speed up openpyxl export
        "rich-argparse==0.5.0",  # colourful help
        "scipy==1.10.1",  # used by others, but also for rankdata
        # -------------------------------------------------------------------------
        # For development:
        # -------------------------------------------------------------------------
        "black==24.3.0",  # auto code formatter
        "flake8==3.8.3",  # code checks
        "pytest==7.1.1",  # automatic testing
    ],
    # Launch scripts:
    entry_points={
        "console_scripts": [
            # Format is 'script=module:function".
            "pdn_project_allocation=pdn_project_allocation.main:main",
            "pdn_project_allocation_run_tests=pdn_project_allocation.run_tests:main",  # noqa
            # noqa
        ],
    },
)
