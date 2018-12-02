from setuptools import setup

setup(
    name='basic_mortgage_calculator',
    description='Generate Mortgage Spreadsheets',
    url='',
    version='0.1.0',
    author='Bent Thumb',
    author_email='redacted@example.com',
    license='MIT',
    packages=['b_mort_calc'],
    install_requires=[
        'xlwt',
    ],
    entry_points = {
        'console_scripts': ['bmc=b_mort_calc.mortgage_calc:main']
    }
)