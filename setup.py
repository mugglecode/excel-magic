import setuptools
from setuptools import setup

setup(
        name='excelmagic',
        version='1.0.0dev1',
        packages=setuptools.find_packages(),
        url='github.com/guo40020/excel-magic',
        license='MIT',
        author='Kelly',
        author_email='',
        description='',
        install_requires=['xlrd', 'XlsxWriter'],
        python_requires='>3.6'
)
