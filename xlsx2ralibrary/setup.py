"""Setup module."""

from setuptools import setup, find_packages


setup(
    name='xlsx2ralibrary',
    version='1.0',
    description='Import data from xlsx file to RaLibrary.',
    url='https://github.com/CVBDL/xlsx2ralibrary',
    author='Patrick Zhong',
    license='MIT',
    packages=find_packages(exclude=['tests']),
    package_data={
        'xlsx2ralibrary': ['assets/certificate.cer']
    },
    install_requires=['openpyxl>=2.4.9', 'requests>=2.18.4']
)
