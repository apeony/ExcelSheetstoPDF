from setuptools import setup, find_packages


setup(
    name="ExcelSheetsToPDF-pkg-unpac", # Replace with your own username
    version="0.1.6",
    author="YEONJU JUNG",
    author_email="unpac.tech@gmail.com",
    description="transform excel file sheets save each pdf file",
    url="https://github.com/apeony/ExcelSheetstoPDF",
    packages=['ExcelSheetsToPDF'],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.9',
)