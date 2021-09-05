from setuptools import setup, find_packages

classifiers = [
    'Development Status :: 4 - Beta',
    'Intended Audience :: Developers',
    'Operating System :: Microsoft :: Windows :: Windows 10',
    'License :: OSI Approved :: MIT License',
    'Programming Language :: Python :: 3'
]

setup(
    name='docxedit',
    version='0.0.1',
    description='Edit Word documents, keep original format.',
    long_description=open('README.txt').read() + '\n\n' + open('CHANGELOG.txt').read(),
    url='https://github.com/henrihapponen/docxedit',
    author='Henri Happonen',
    author_email='henkka.happonen@gmail.com',
    license='MIT',
    classifiers=classifiers,
    keywords=['docx', 'python-docx', 'docxedit', 'docx-edit'],
    packages=find_packages,
    install_requires='python-docx'
)
