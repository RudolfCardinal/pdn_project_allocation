#!/usr/bin/env/python

"""
setup.py
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
        'lxml==4.6.3',  # Will speed up openpyxl export
    ],

    # Launch scripts:
    entry_points={
        'console_scripts': [
            # Format is 'script=module:function".
            'pdn_project_allocation=pdn_project_allocation.pdn_project_allocation:main',  # noqa
            'pdn_project_allocation_run_tests=pdn_project_allocation.run_tests:main',  # noqa
            # noqa
        ],
    },
)
