<?php

namespace Brezgalov\XlsBuilder;

use Dompdf\Dompdf;
use PhpOffice\PhpSpreadsheet\IOFactory;
use yii\base\Model;
use yii\helpers\ArrayHelper;
use yii\web\View;

/**
 * Class BaseTemplateBuilder
 * @package common\models\forms\XlsTemplates
 */
abstract class BaseTemplateBuilder extends Model
{
    /**
     * @var string префикс для временных файлов
     */
    public $filePrefix = 'xls_builder_tmp_file';

    /**
     * @var string название шаблона для вывода в заголовке письма
     */
    public $mailTitle = 'Базовый';

    /**
     * @var string путь к view для письма
     */
    public $mailView;

    /**
     * @var string
     */
    public $defaultMailSubject = 'Пересылка докуметов';

    /**
     * @var string css для правки стилей в pdf
     */
    public $cssFix = '';

    /**
     * @var array
     */
    protected $props = [];

    /**
     * Тут мы будем держать методы для мутирования входящих данных
     * @return callable[]
     */
    public function getMutations()
    {
        return [];
    }

    /**
     * Массив вида ['test_1' => ['A1', 'A2'], 'test_2' => ['B3']]
     * @return array
     */
    public abstract function getPropsMap();

    /**
     * Полный путь до файла шаблона
     * @return string
     */
    public abstract function getTemplatePath();

    /**
     * @param array $props
     */
    public function setProps(array $props)
    {
        $this->props = array_merge($this->props, $props);
    }

    /**
     * @return array
     */
    public function getTemplateProps()
    {
        return $this->props;
    }

    /**
     * @return \PhpOffice\PhpSpreadsheet\Spreadsheet|bool
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function buildFile()
    {
        $filePath = $this->getTemplatePath();
        if (!file_exists($filePath)) {
            $this->addError('file', 'Не удается найти файл шаблона');
            return false;
        }

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load($filePath);

        $page = $spreadsheet->getActiveSheet();
        $propsMap = $this->getPropsMap();

        $mutations = $this->getMutations();
        foreach ($this->getTemplateProps() as $key => $value) {
            $map = ArrayHelper::getValue($propsMap, $key, []);

            if (array_key_exists($key, $mutations)) {
                $value = call_user_func($mutations[$key], $value);
            }

            foreach ($map as $pos) {
                $page->setCellValue($pos, $value);
            }
        }

        return $spreadsheet;
    }

    /**
     * @return string
     */
    public function generateFileName()
    {
        return uniqid($this->filePrefix, true);
    }

    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
     * @return bool|string
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function saveFile($spreadsheet)
    {
        if (!$spreadsheet) {
            if (!$this->hasErrors()) {
                $this->addError('file', 'Не известная ошибка');
            }
            return false;
        }

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $tmpFile = \Yii::getAlias('@runtime/') . $this->generateFileName();
        $filePath = $tmpFile . '.xlsx';
        $writer->save($filePath);

        $as_pdf = ArrayHelper::getValue($this->props, 'as_pdf', false);
        if ($as_pdf) {
            $writer = IOFactory::createWriter($spreadsheet, 'Html');
            $writer->save($tmpFile . '.html');
            unlink($filePath);

            $filePath = $tmpFile . '.pdf';

            $dompdf = new Dompdf();
            $html = file_get_contents($tmpFile . '.html');

            $styleIndex = strpos($html, "</head>");
            $string = "\n<style>{$this->cssFix}</style>\n";
            $html = substr_replace($html, $string, $styleIndex, 0);

            unlink($tmpFile . '.html');
            $dompdf->loadHtml($html);
            $dompdf->render();

            if (!file_put_contents($filePath, $dompdf->output())) {
                $this->addError('file', 'Не удается сохранить файл');
                return false;
            }
        }

        return $filePath;
    }

    /**
     * @param $filePath
     * @param $email
     */
    public function sendEmail($filePath, $email)
    {
        if (strpos($email, ',') !== false) {
            $email = explode(',', $email);
        }

        $mailSetup  = $this->getMailSetup();
        $mailTitle  = ArrayHelper::getValue($mailSetup, 'title', "Файл шаблона {$this->mailTitle}");
        $mailBody   = (new View())->render($this->mailView);
        \Yii::$app->mailer->compose()
            ->setFrom(\Yii::$app->params['noReplyEmail'])
            ->setTo(\Yii::$app->params['noReplyEmail'])
            ->setBcc($email)
            ->setSubject($mailTitle)
            ->setHtmlBody($mailBody)
            ->attach($filePath)
            ->send();
    }

    /**
     * @return array
     */
    public function getMailSetup()
    {
        return [
            'title' => $this->defaultMailSubject,
            'text'  => '',
        ];
    }
}