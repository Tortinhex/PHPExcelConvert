<?php
include './Motor.php';

$html = $_POST['html'];

$conversor = new Motor($html);
$conversor->setNameDocument("AST");

$conversor->setNameAba("Pasta de trabalho");
$conversor->setColorFundo("#FFFFFF");
$conversor->setColorHeader("AAAAAA");
$conversor->setColorsStrip("#FFFFFF", "EEEEEE");
$conversor->setFontSizeTitle(16);
$conversor->setFontSizeBody(12);
$conversor->setFontFamily("verdana");
$conversor->setDefaultPositionX(3);
$conversor->setDefaultPositionY(3);

$conversor->gerarXls();

