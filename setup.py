from setuptools import setup, find_packages

setup(
    name='goup-setup',
    version='1.0.0',
    description='Uma descrição do meu pacote',
    packages=find_packages(),
    install_requires=[
        'requests',
        'openpyxl',
        'streamlit',
    ],
)
