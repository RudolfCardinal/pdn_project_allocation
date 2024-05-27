..  README.rst

..  Copyright (C) 2019-2021 Rudolf Cardinal (rudolf@pobox.com).
    .
    This file is part of pdn_project_allocation.
    .
    This is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    .
    This software is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
    GNU General Public License for more details.
    .
    You should have received a copy of the GNU General Public License
    along with this software. If not, see <http://www.gnu.org/licenses/>.

.. _Meld: https://meldmerge.org/


.. Code style:
.. image:: https://img.shields.io/badge/code%20style-black-000000.svg
    :target: https://github.com/psf/black


:ref: wininst
Windows Installation Guide
==========================

This software can be installed on Windows using the instructions in the ``README.rst``
file but some additional software (e.g. git) may need to be installed.  The instructions below are designed to help those with less experience who wish to deploy it on Windows without these additional installations.  These instructions avoid the need to install
anything other than Python and this software itself.

You will need to work with the command line in Windows.  If you are unfamiliar with this, you can find various introductions on the web (for example: https://www.makeuseof.com/tag/a-beginners-guide-to-the-windows-command-line/)

Installation
------------
- If you don't already have it installed, download and install Python (https://www.python.org/downloads/windows/).  You can use any version between 3.8 and 3.11.9 inclusive (the latter is near the top of the page).  Get the correct installer for your system (this is probably the Windows installer 64-bit file if you are running Windows 10 or 11).  Please note that Python 3.12.x is not currently compatible with this software.

- Open a Windows command prompt (type `cmd.exe` in the Windows search box)

- Create a directory in which you wish to install the software e.g.:

  .. code-block:: console

    mkdir pdn_project_allocation

  and navigate into it:

  .. code-block:: console

    cd pdn_project_allocation

- Create a Python 3 virtual environment e.g.:

  .. code-block:: console

    py -m venv pdn_venv

  and activate it:

  .. code-block:: console

    pdn_venv\Scripts\activate

- Install pdn_project_allocation direct from github:

  .. code-block:: console

    pip install https://github.com/RudolfCardinal/pdn_project_allocation/archive/refs/heads/master.zip

You should now be able to run the program. Try:

.. code-block:: bash

    pdn_project_allocation --help

To run some automated tests, change into a directory where you're happy to
stash some output files and run

.. code-block:: bash

    pdn_project_allocation_run_tests

This produces solutions to match the test data in the
``pdn_project_allocation/testdata`` directory.

Don't forget that to use the software the virtual environment must have been activated.  If you close your command prompt and come back another time to use it again, just use the activation command shown above (``pdn_venv\Scripts\activate``), in the directory in which you installed the software.
