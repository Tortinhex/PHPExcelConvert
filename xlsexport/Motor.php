<?php

include 'PHPExcel.php';

class Motor {

    private $objPHPExcel;
    private $html;
    private $nomeArquivo = "arquivo.xls";
    private $nomeAba = "arquivo";
    private $colorStrip1;
    private $colorStrip2;
    private $colorTitle;
    private $fontSizeBody = "11";
    private $fontSizeTitle = "20";
    private $posX = "2";
    private $posY = "1";
    private $fontFamily = "arial";
    private $alignTitle = "left";
    private $alfa;
    private $contColumns = 0;
    private $contRows = 0;
    private $striped = false;

    function __construct($html) {
        //Starta a API
        $this->iniciarApi($html);
        $this->html = $html;
    }

    /*
     * Inicializacao da API
     */

    private function iniciarApi($html) {

        $this->objPHPExcel = new PHPExcel();
        $this->setAlfa();
    }

    /**
     * Define o nome do documento
     */
    public function setNameDocument($name) {
        if ($name > 1) {
            if (!eregi(".xls", $name)) {
                $this->nomeArquivo = $name . ".xls";
            } else {
                $this->nomeArquivo = $name;
            }
        }
    }

    /**
     * Troca o nome da aba de trabalho
     */
    public function setNameAba($name) {
        $this->nomeAba = $name;
    }

    /**
     * Remove o '#' caso a string tenha
     */
    private function removeHash($string) {
        if (eregi("#", $string)) {
            return substr($string, 1);
        } else {
            return $string;
        }
    }

    /**
     * Define a cor de fundo do documento
     */
    public function setColorFundo($cor) {
        $this->objPHPExcel->getActiveSheet()->getStyle('A:Z')->applyFromArray(
                array(
                    'fill' => array(
                        'type' => PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('rgb' => $this->removeHash($cor)),
                    ),
                )
        );
    }

    /**
     * Define as cores que irao variar
     * de linha em linha, criando o efeito de strip
     */
    public function setColorsStrip($cor1, $cor2) {
        $this->setStriped(true);
        $this->colorStrip1 = $this->removeHash($cor1);
        $this->colorStrip2 = $this->removeHash($cor2);
    }

    /**
     * Define a cor do titulo
     */
    public function setColorHeader($cor) {
        $this->colorTitle = $this->removeHash($cor);
    }

    /**
     * Define a fonte do body
     */
    public function setFontSizeBody($size) {
        $this->fontSizeBody = $size;
    }

    /**
     * Define a fonte do titulo
     */
    public function setFontSizeTitle($size) {
        $this->fontSizeTitle = $size;
    }

    /**
     * Define a fonte que vai ser usada na tabela
     */
    public function setFontFamily($font) {
        $this->fontFamily = $font;
    }

    /**
     * Define a posicao inicial da tabela na celula em X
     */
    public function setDefaultPositionX($pos) {
        $this->posX = $pos;
    }

    /**
     * Define a posicao inicial da tabela na celula em Y
     */
    public function setDefaultPositionY($pos) {
        $this->posY = $pos - 1;
    }

    /**
     * Define se o xls gerado contera o modelo
     * de formatacao striped
     */
    public function setStriped($condition) {
        $this->striped = $condition;
    }

    /**
     * Gera todos os estilos da tabela
     */
    public function generateStyles() {
        $strip = false;


        for ($i = $this->posX; $i < ($this->contRows + $this->posX); $i++) {
            if ($i == $this->posX) { /* Se for a primeira linha, sera considerada header */
                $this->objPHPExcel->getActiveSheet()->getRowDimension($i)->setRowHeight($this->fontSizeTitle + 5);

                $style = array();

                $style['font'] = array(
                    'name' => $this->fontFamily,
                    'size' => $this->fontSizeTitle,
                    'bold' => true
                );

                if (strlen($this->colorTitle) > 2) {
                    $style['fill'] = array(
                        'type' => PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('rgb' => $this->colorTitle),
                    );
                }

                $style['alignment'] = array(
                    'horizontal' => $this->setAlignTitle($this->alignTitle)
                );

                $this->objPHPExcel->getActiveSheet()->getStyle($this->alfa[$this->posY + 1] . $i . ':' . $this->alfa[$this->contColumns + $this->posY] . $i)->applyFromArray($style);
            } else {
                $this->objPHPExcel->getActiveSheet()->getRowDimension($i)->setRowHeight($this->fontSizeBody + 5);
                $style = array();
                $font = array('font' => array(
                        'name' => $this->fontFamily,
                        'size' => $this->fontSizeBody
                ));
                $this->objPHPExcel->getActiveSheet()->getStyle($this->alfa[$this->posY + 1] . $i . ':' . $this->alfa[$this->contColumns + $this->posY] . $i)->applyFromArray($font);

                if ($this->striped) {

                    if ($strip) {

                        $style['fill'] = array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => $this->colorStrip1),
                        );

                        $this->objPHPExcel->getActiveSheet()->getStyle($this->alfa[$this->posY + 1] . $i . ':' . $this->alfa[$this->contColumns + $this->posY] . $i)->applyFromArray($style);
                        $strip = false;
                    } else {
                        //Se for false, pinta a linha de outra cor

                        $style['fill'] = array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => $this->colorStrip2),
                        );

                        $this->objPHPExcel->getActiveSheet()->getStyle($this->alfa[$this->posY + 1] . $i . ':' . $this->alfa[$this->contColumns + $this->posY] . $i)->applyFromArray($style);
                        $strip = true;
                    }
                }
            }
        }
    }

    /**
     * Converte o html e escreve em formato xls
     */
    private function writeTable($html) {

        $auxX = $this->posX;
        $auxY = $this->posY;
        $DOM = new DOMDocument;
        @$DOM->loadHTML($html);
        $items = $DOM->getElementsByTagName("tr");

        foreach ($items as $node) {
            $cont = 0;
            foreach ($node->childNodes as $element) {
                $this->objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($auxY, $auxX, $element->nodeValue);
                $auxY++;
                $cont++;
            }

            if ($cont > $this->contColumns) {
                $this->contColumns = $cont;
            }

            $auxX++;

            $auxY = $this->posY;
            $this->contRows++;
        }
    }

    /**
     * Define o alinhamento do titulo
     */
    public function setAlignTitle($align) {
        switch ($align) {
            case "left":
                $this->alignTitle = array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                break;
            case "center":
                $this->alignTitle = array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                break;
            default:
                $this->alignTitle = array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
                break;
        }
    }

    /**
     * Gera o arquivo XLS
     */
    public function gerarXls() {
        $this->writeTable($this->html);
        $this->objPHPExcel->getActiveSheet()->setTitle($this->nomeAba);
        $this->generateStyles();
        $this->adjustAutoSize();

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $this->nomeArquivo . '"');
        header('Cache-Control: max-age=0');
        header('Cache-Control: max-age=1');

        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');

        $objWriter->save('php://output');
    }

    /**
     * Hack de auto ajuste (width) das tabelas
     */
    private function adjustAutoSize() {
        foreach ($this->alfa as $item) {
            $this->objPHPExcel->getActiveSheet()->getColumnDimension($item)->setAutoSize(true);
        }
    }

    /**
     * Tabela de conversao de numeros para letras
     */
    private function setAlfa() {
        $this->alfa = array(
            1 => "A",
            2 => "B",
            3 => "C",
            4 => "D",
            5 => "E",
            6 => "F",
            8 => "G",
            9 => "H",
            10 => "I",
            11 => "J",
            12 => "K",
            13 => "L",
            14 => "M",
            15 => "N",
            16 => "O",
            17 => "P",
            18 => "Q",
            19 => "R",
            20 => "S",
            21 => "T",
            22 => "U",
            23 => "V",
            24 => "W",
            25 => "X",
            26 => "Y",
            27 => "Z"
        );
    }

}
