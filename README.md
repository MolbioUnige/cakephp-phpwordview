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



Two different ways to render.

### PHPWord TemplateProcessor

To be used if you need to render simple documents. Serialize the data you want to insert into the template.

The templates are expected to be present in `src/Templates/Model/docx/`, and having `.docx` filename extension. Template files are just plain word document, not Word model documents.

[Create a template document](https://phpword.readthedocs.io/en/latest/templates-processing.html)

```php
public function imprimatur()
{
    $imprimatur = '667';
    $this->viewBuilder()
         ->setClassName('PhpWordView.PhpWord');
         
    $data = [
        'title' => 'Beati Pauperes Spiritu,
        'imprimatur' => $imprimatur,
    ];
    $_serialize = 'data';
    $this->set(compact('data', '_serialize'));
}
```

### PHPWord full api

To be used when a simple search-replace is not sufficient.

The view files are expected to be present in `src/Templates/Model/docx/`, but having `.ctp` filename extension. Inside your view files you will have access to the PHPWord library with `$this->PhpWord`. Check the [PHPWord](https://github.com/PHPOffice/PHPWord) documentation on how to use PHPWord.

Don't set the `_serialize` variable.

```php
public function imprimatur($id)
{
    $this->viewBuilder()
         ->setClassName('PhpWordView.PhpWord');
    $imprimatur = $this->Imprimaturs->get($id);
    $this->set(compact('imprimatur'));
}
```

### Downloaded filename

The downloaded filename can be specified with the viewBuilder options, default is the action name.

```php
    $this->viewBuilder()
         ->setClassName('PhpWordView.PhpWord')
         ->options([
             'wordConfig' => [
                 'filename' => 'imprimatur-' . $id,
             ]
         ]);

```

## Troubleshooting

A `Could not close zip file /tmp/PhpWord5gJC0Z` Exception might mean that there is a problem with the template.docx file have incorrect line ending. This can be solved by specifying docx files as binary in the `.gitattributes` file.
