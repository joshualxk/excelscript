from setuptools import setup

setup(
    name='excelscript',
    version='1.0',
    description='excel拆表',
    packages=['excelscript'],
    package_dir={'excelscript': 'excelscript'},
    package_data={'excelscript': ['Shin-chan.png', 'config.yml']},
    scripts=[
        'scripts/excelsc.bat',
    ]
)
