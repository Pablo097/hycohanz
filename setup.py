#!/usr/bin/env python

from setuptools import setup

setup(name='hycohanz',
      description='Interact with ANSYS HFSS via the HFSS Windows COM API.',
      author='Matthew Radway',
      author_email='mradway@gmail.com',
      version='0.0.2pre',
      packages=['hycohanz'],
      install_requires=['pywin32', 'quantiphy']
      )
