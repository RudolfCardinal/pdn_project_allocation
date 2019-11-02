#!/usr/bin/env/python
# setup.py

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
        'License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)', # noqa
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Topic :: Education'
    ],
    packages=find_packages(),
    install_requires=[
        'cardinal_pythonlib >= 1.0.71',
        'mip >= 1.5.3',
    ],
    entry_points={
        'console_scripts': [
            # Format is 'script=module:function".
            'pdn_project_allocation=pdn_project_allocation.pdn_project_allocation:main',  # noqa
        ],
    },
)
