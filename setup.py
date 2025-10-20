from setuptools import setup, find_packages

setup(
    name="MarkdownToPPTX",
    version="0.1.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        # Add your dependencies here
    ],
    author="mykLabs",
    author_email="mikkel03@gmail.com",
    description="A short description of the project",
    python_requires=">=3.10.6",
)
