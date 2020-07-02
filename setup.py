import setuptools
from setuptools import setup

setup(
        name='excelmagic2',
        version='2.0.0',
        packages=setuptools.find_packages(),
        url='https://github.com/mugglecode/excel-magic/tree/2.0',
        license='MIT',
        author='Kelly',
        author_email='',
        description='',
        install_requires=['openpyxl', 'XlsxWriter', 'Pillow'],
        python_requires='>3.7.2'
)
