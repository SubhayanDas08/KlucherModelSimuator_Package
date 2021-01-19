from setuptools import setup, find_packages

with open("README.md", "r") as fh:
  long_description = fh.read()

classifiers = [
  'Development Status :: 5 - Production/Stable',
  'Intended Audience :: Education',
  'Operating System :: Microsoft :: Windows :: Windows 10',
  'License :: OSI Approved :: MIT License',
  'Programming Language :: Python :: 3'
]
 
setup(
  name='KlucherModelSDas',
  version='1.0.1',
  description='Simulator for Klucher Model',
  long_description=open('README.md').read() + '\n\n' + open('CHANGELOG.txt').read(),
  url='https://github.com/SubhayanDas08/KlucherModelSimuator_Package',   
  author='Subhayan Das',
  author_email='nips.subhayan@gmail.com',
  license='MIT', 
  classifiers=classifiers,
  keywords=['klucher', 'solar', 'simulator'], 
  packages=find_packages(),
  install_requires=[
    'XlsxWriter ~= 1.3.7',
  ],
  long_description_content_type="text/markdown",
  extras_require = {
    "dev": [
      "pytest>=3.7",
    ]
  }
)