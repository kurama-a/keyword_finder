from setuptools import setup, find_packages

setup(
    name='extraction_script',
    version='1.0',
    packages=find_packages(),
    install_requires=[
        'pandas',
        'python-docx',
        'python-pptx',
        'PyPDF2',
        'Pillow',
        'pytesseract',
        'pdf2image'
    ],
    entry_points={
        'console_scripts': [
            'extraction_script=keyworld_finder:search_keywords_in_files',
        ],
    },
    include_package_data=True,
    description='Script d\'extraction de texte et d\'images à partir de différents fichiers.',
    author='Paul ARNAUD',
    url='https://github.com/kurama-a/keyword_finder',
    classifiers=[
        'Programming Language :: Python :: 3'
    ],
)
