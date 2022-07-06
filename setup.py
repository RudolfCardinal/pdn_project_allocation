#!/usr/bin/env/python

"""
setup.py

===============================================================================

    Copyright (C) 2019-2021 Rudolf Cardinal (rudolf@pobox.com).

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

setup(
    name='pdn_project_allocation',
    version='0.0.1',
    description='Allocate students to projects',

    url='https://github.com/RudolfCardinal/pdn_project_allocation.git',
    author='Rudolf Cardinal',
    author_email='rudolf@pobox.com',

    license='GNU General Public License v3 or later (GPLv3+)',

    # See https://pypi.org/classifiers/
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Education',
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)',  # noqa
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Topic :: Education'
    ],

    # Python code:
    packages=find_packages(),

    # Static files:
    # https://stackoverflow.com/questions/11848030/how-include-static-files-to-setuptools-python-package
    package_data={
        'pdn_project_allocation.testdata': ['*'],
    },

    # Requirements:
    install_requires=[
        'cardinal_pythonlib==1.1.7',
        'mip==1.13.0',
        'matching==1.4',
        'openpyxl==3.0.9',
        'lxml==4.9.1',  # Will speed up openpyxl export
    ],

    # Launch scripts:
    entry_points={
        'console_scripts': [
            # Format is 'script=module:function".
            'pdn_project_allocation=pdn_project_allocation.main:main',
            'pdn_project_allocation_run_tests=pdn_project_allocation.run_tests:main',  # noqa
            # noqa
        ],
    },
)
