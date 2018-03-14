# Remove password on viewing VBA source code

>Remove password on viewing VBA source code on MS Office documents

## Description

This PHP script will allow you to reset and then remove the password needed for viewing VBA source code on MS Office documents.

## Table of Contents

- [Author](#author)
- [Install](#install)
- [Usage](#usage)
- [License](#license)

## Author

This script has been coded by [AyrA](https://github.com/AyrA).
Original source code can be found [here](https://github.com/AyrA/ExcelStuff#exceldecrypt)

## Install

Get a copy of the index.php script, save it on your localhost server and access to the script from within your browser (like `http://localhost/vba_decrypt/index.php`)

## Usage

1. Access to http://localhost/vba_decrypt/index.php or any other location where you've saved the script.
2. Follow instructions displayed on screen i.e.
    a. Upload the Office protected document
    b. Click on the `Decrypt VBA` button
    c. Open the downloaded file, the "almost" unprotected copy of your file
    d. Ignore the alerts by answering `Yes` to `Do you want to continue to open the file?`
    e. Go to the protection tab and type a new password (test / test f.i.) and click on OK. This will in fact reset the password.
    f. Go once more to the protection tab and just uncheck the `Lock project for viewing`

Tadaaaa, the VBA password protection has been resetted and removed.

## License

[MIT](LICENSE)