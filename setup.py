from setuptools import setup, find_packages

classifiers = [
    'Development Status :: 5 - Production/Stable',
    'Intended Audience :: Developers',
    'Operating System :: Microsoft :: Windows :: Windows 10',
    'License :: OSI Approved :: MIT License',
    'Programming Language :: Python :: 3'
]

setup(
    name='docx-edit',
    version='0.0.1',
    description='Edit Word documents, keep original format.',
    long_description=open('README.txt').read() + '\n\n' + open('CHANGELOG.txt').read(),
    url='',
    author='Henri Happonen',
    author_email='henkka.happonen@gmail.com',
    license='MIT',
    classifiers=classifiers,
    keywords='docx',
    packages=find_packages,
    install_requires='docx'
)
