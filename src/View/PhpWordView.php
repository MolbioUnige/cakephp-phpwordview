<?php
namespace PhpWordView\View;

use Cake\Core\Configure;
use Cake\Core\Exception\Exception;
use Cake\Event\EventManager;
use Cake\Http\Response;
use Cake\Http\ServerRequest;
use Cake\View\View;
use PhpOffice\PhpWord\TemplateProcessor;

/**
 * A view class to render PhpWord views.
 *
 * By setting the 'serialize' view builder option, you can specify a view variable
 * that will replace to corresponding placeholder in the template file.
 *
 * This plugin uses only the templateRender functionnality of PhpWord.
 */
class PhpWordView extends View
{
    /**
     * PhpWord views are always in docx.
     *
     * @var string
     */
    protected $subDir = 'docx';

    /**
     * Word layouts/models are located in the docx sub directory of `Layouts/`
     *
     * @var string
     */
    protected $layoutPath = 'docx';

    /**
     * List of pdf configs collected from the associated controller.
     *
     * @var array
     */
    public $wordConfig = [];

    /**
     * File extension. Defaults to PhpWord's template ".docx".
     *
     * @var string
     */
    protected $_ext = '.docx';

    /**
     * Constructor
     *
     * @param \Cake\Http\ServerRequest $request Request instance
     * @param \Cake\Http\Response $response Response instance
     * @param \Cake\Event\EventManager $eventManager Event manager instance
     * @param array $viewOptions View options.
     *
     * @throws \Cake\Core\Exception\Exception
     */
    public function __construct(
        ServerRequest $request = null,
        Response $response = null,
        EventManager $eventManager = null,
        array $viewOptions = []
    ) {
        $this->_passedVars[] = 'wordConfig';

        parent::__construct($request, $response, $eventManager, $viewOptions);

        $this->response = $this->response->withType('docx');

        if (isset($viewOptions['templatePath']) && $viewOptions['templatePath'] == 'Error') {
            $this->subDir = null;
            $this->layoutPath = null;
            $this->response = $this->response->withType('html');

            return;
        }
    }

    /**
     * Render a PhpWord view.
     *
     * Uses the special '_serialize' parameter to convert a set of
     * view variables into a PhpWord response. Makes generating simple
     * PhpWord responses very easy. If you omit the '_serialize' parameter,
     * templates vairables will not be substituted.
     *
     * @param string|null $view   The view being rendered.
     * @param string|null $layout The layout being rendered.
     *
     * @return string The rendered view.
     */
    public function render($view = null, $layout = null)
    {
        $filename = isset($this->wordConfig['filename']) && is_string($this->wordConfig['filename'])
            ? $this->wordConfig['filename']
            : $this->getTemplate();
        $this->response = $this->response->withDownload($filename);

        $serialize = $this->viewVars['_serialize'];
        ob_start();
        $templateProcessor = new TemplateProcessor($this->_getViewFileName());
        foreach ((array)$serialize as $viewVar) {
            if (is_scalar($this->viewVars[$viewVar])) {
                throw new Exception("'" . $viewVar . "' is not an array or iteratable object.");
            }

            foreach ($this->viewVars[$viewVar] as $key => $value) {
                $templateProcessor->setValue($key, $value);
            }
        }
        $templateProcessor->saveAs('php://output');

        return ob_get_clean();
    }
}
