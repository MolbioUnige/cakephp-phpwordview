<?php
namespace PhpWordView\Test\TestCase\View;

use Cake\Http\Response;
use Cake\Http\ServerRequest as Request;
use Cake\TestSuite\TestCase;
use PhpWordView\View\PhpWordView;

/**
 * PhpWordViewTest
 */
class PhpWordViewTest extends TestCase
{
    public function setUp()
    {
        // Time::setToStringFormat('yyyy-MM-dd HH:mm:ss');

        $this->request = new Request();
        $this->response = new Response();

        $this->view = new PhpWordView($this->request, $this->response);
        $this->view->setTemplatePath('Imprimaturs');
        $this->view->setSubDir('docx');
        $this->view->setTemplate('blank');
    }

    /**
     * TestRenderDefault
     */
    public function testRenderDefault()
    {
        $this->view->set(['_serialize' => []]);
        $this->view->render('view');
        $this->assertSame(
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            $this->view->getResponse()->getType()
        );
        $this->assertContains(
            'filename="blank"',
            $this->view->getResponse()->getHeaderLine('Content-Disposition')
        );
    }

    /**
     * TestRenderCustomFilename
     */
    public function testRenderCustomFilename()
    {
        $this->view->wordConfig['filename'] = 'custom';
        $this->view->set(['_serialize' => []]);
        $this->view->render(false);
        $this->assertContains(
            'filename="custom"',
            $this->view->getResponse()->getHeaderLine('Content-Disposition')
        );
    }
}
