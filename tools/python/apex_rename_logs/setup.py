from setuptools import setup, find_packages

setup(
    name="apex_rename_logs",
    version="1.0.0",
    description="Outil CLI pour renommer les fichiers journaux obsolètes en .DEPRECATED (compatible WSL/Windows)",
    author="APEX Framework Team",
    packages=find_packages(),
    py_modules=["rename_logs"],
    entry_points={
        "console_scripts": [
            "apex-rename-logs=rename_logs:main"
        ]
    },
    python_requires=">=3.6",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    install_requires=[],
    include_package_data=True,
    zip_safe=False,
) 