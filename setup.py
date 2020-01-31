#
# Copyright (c) 2016, Prometheus Research, LLC
#


from setuptools import setup, find_packages


setup(
    name='htsql_excel',
    version='0.1.4',
    description='An HTSQL extension that adds basic Excel support.',
    long_description=open('README.rst', 'r').read(),
    keywords='htsql extension excel xls xlsx',
    author='Prometheus Research, LLC',
    author_email='contact@prometheusresearch.com',
    license='Apache-2.0',
    classifiers=[
        'Programming Language :: Python :: 2.7',
        'License :: OSI Approved :: Apache Software License',
    ],
    url='https://github.com/prometheusresearch/htsql_excel',
    package_dir={'': 'src'},
    packages=find_packages('src'),
    zip_safe=True,
    include_package_data=True,
    entry_points={
        'htsql.addons': [
            'htsql_excel = htsql_excel:ExcelAddon',
        ],
    },
    install_requires=[
        'HTSQL>=2.3,<3',
        'xlwt>=1,<2',
        'openpyxl>=2,<3',
    ],
)

