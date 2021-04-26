from setuptools import setup, find_packages

setup(
    name="sendmail",
    version="1.0.0",
    description="Send attachment email with use command line.",
    author="norik00",
    packages=find_packages(),
    install_requires=["PyYAML"],
    entry_points={
        "console_scripts": [
            "sendmail=sendmail.sendMail:main",
        ]
    },
    classifiers=[
        "Programing Language :: Python :: 3.9.0"
    ]
)