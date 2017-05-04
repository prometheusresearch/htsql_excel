#
# Copyright (c) 2016, Prometheus Research, LLC
#


from setuptools import setup, find_packages


setup(
    name='htsql_excel',
    version='0.1.1',
    description='An HTSQL extension that adds basic Excel support.',
    long_description=open('README.rst', 'r').read(),
    keywords='htsql extension excel xls xlsx',
    author='Prometheus Research, LLC',
    author_email='contact@prometheusresearch.com',
    license='AGPLv3',
    classifiers=[
        'Programming Language :: Python :: 2.7',
        'License :: OSI Approved :: GNU Affero General Public License v3',
    ],
    url='https://bitbucket.org/prometheus/htsql_excel',
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

