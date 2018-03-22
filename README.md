# DocxToWhatsapp
[![made-with-python](https://img.shields.io/badge/Made%20with-Python-1f425f.svg)](https://www.python.org/) [![GPLv3 license](https://img.shields.io/badge/License-GPLv3-blue.svg)](http://perso.crans.org/besson/LICENSE.html)  [![Maintenance](https://img.shields.io/badge/Maintained%3F-no-red.svg)](https://bitbucket.org/lbesson/ansi-colors) 
A small python script that will open a `.docx` file and do some modifications that will allow me to copy &amp; paste in Whatsapp.

![Screenshot](screenshot.png?raw=true "Screenshot")

### Prerequisites

You need :
- python 3.6
- pyinstaller (if you want to make an exe)
  - `pip install pyinstaller` 

### Release History
Actual version : 3
1. Basic
2. GUI + copy&paste
3. Allow to create an exe file 

### Dependency
- python : pyperclip
- Create an executable with pyinstaller
  - `pyinstaller .\DocxToWhatsapp.py --clean -w`

## Author
- **Yassine Gangat** - [gyassine](https://github.com/gyassine)

## Contributing
1. Fork it (<https://github.com/gyassine/DocxToWhatsapp/fork>)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

## Acknowledgments
* Inspired by <https://gist.github.com/etienned/7539105> which only extract paragraph from word file

## License

The content of this project itself is licensed under the [Creative Commons Attribution 3.0 license](http://creativecommons.org/licenses/by/3.0/us/deed.en_US), and the underlying source code used to format and display that content is licensed under the [GNU General Public License v3.0](https://choosealicense.com/licenses/gpl-3.0/).