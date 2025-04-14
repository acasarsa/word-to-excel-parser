# automation-project/setup.py
from setuptools import setup, find_packages

setup(
    name='automation_app',
    version='0.1',
    packages=find_packages(),
    install_requires=[
        'python-docx',
        'openpyxl',
    ],
    entry_points={
        'console_scripts': [
            'automation-app-install=automation_app.setup_script:install',
        ],
    },
)
