from setuptools import setup, find_packages

with open('README.md', encoding='utf8') as readme_file:
    readme = readme_file.read()

with open('requirements.txt') as f:
    requirements = f.read().splitlines()

setup(
    name='pyxcel',
    version='1.0.0',
    description='A Python Library for Excel Workbooks.',
    long_description=readme,
    long_description_content_type='text/markdown',
    author='Daniel Kelshaw',
    author_email='daniel.j.kelshaw@gmail.com',
    url='https://github.com/Design4Demise/PyExcelAnalysis',
    packages=find_packages(),
    install_requires=requirements,
    license='MIT License'
)
