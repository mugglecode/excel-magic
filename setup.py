import setuptools
from setuptools import setup

setup(
        name='excelmagic',
        version='1.3.2',
        packages=setuptools.find_packages(),
        url='https://github.com/guo40020/excel-magic',
        license='MIT',
        author='Kelly',
        author_email='',
        description='',
        install_requires=['xlrd', 'XlsxWriter', 'Pillow'],
        python_requires='>3.6'
)
