from setuptools import setup

setup(
    name="gomind_excel",
    python_requires=">=3.6",
    version="1.0.0",
    description="GoMind excel functions",
    url="https://github.com/GrupoDomini/gomind_excel.git",
    author="JeffersonCarvalhoGD",
    author_email="jefferson.carvalho@grupodomini.com",
    license="unlicense",
    packages=["gomind_excel"],
    zip_safe=False,
    install_requires=[
        "pywin32",
    ],
)
