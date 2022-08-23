# PhpWordView plugin for CakePHP

This plugin renders Word documents using the [PhpWord](https://github.com/PHPOffice/PHPWord) library. This plugin uses the `TemplateProcessor` of PhpWord.

## Installation

You can install this plugin into your CakePHP application using [composer](https://getcomposer.org).

The recommended way to install composer packages is:

```
composer require molbio-unige/php-word-view
```

### Enable plugin

Load the plugin by running command

```
bin/cake plugin load PhpWordView
```

## Usage

The templates are expected to be present in `src/Templates/Model/docx/`. Template files are just plain word document, not Word model documents.

[Create a template document](https://phpword.readthedocs.io/en/latest/templates-processing.html)

```php
public function imprimatur()
{
    $imprimatur = '667';
    $this->viewBuilder()
         ->setClassName('PhpWordView.PhpWord')
         ->options([
             'wordConfig' => [
                 'filename' => 'imprimatur-' . $imprimatur,
             ]
         ]);
    $data = [
        'title' => 'Beati Pauperes Spiritu,
        'imprimatur' => $imprimatur,
    ];
    $_serialize = 'data';
    $this->set(compact('data', '_serialize'));
}
```

The downloaded filename can be specified with the viewBuilder options.

## Troubleshooting

A `Could not close zip file /tmp/PhpWord5gJC0Z` Exception might mean that there is a problem with the template.docx file have incorrect line ending. This can be solved by specifying docx files as binary in the `.gitattributes` file. 
