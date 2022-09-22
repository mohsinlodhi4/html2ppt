<?php

require __DIR__.'/vendor/autoload.php';

\PhpOffice\PhpPresentation\Autoloader::register();
\PhpOffice\Common\Autoloader::register();

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Fill;

// echo "<pre>";
// print_r( get_class_methods(PhpPresentation::class) );



$data = [
    [ // slide1
        ['logo'=>  './logo.png'],
        ['h1'=> 'Parabéns, sua avaliação está pronta! '],
        ['h2' => ' Perfil'],
        ['p2' => ' Empresa Teste Ltda - Santa Cruz do Sul, RS
Fundada em 2020'],
        ['h3'=>'Operações'],
        ['p3'=>['text'=>'Estágio da empresa: Crescimento inicial 
Número de sócios: 2 
Experiência da equipe: Mais de 10 anos
Propriedades intelectuais: Marca registrada, Certificações, Domínio / Website / Softwares', 'lines'=>4]],
        ['h4'=>'Oportunidade'],
        ['p4'=>'Média concorrência
Negócio não escalável'],
        ['table'=> [
                ['SETOR', 'EMPRESAS', 'D/E'],
                ['Software (Internet)', '40', '15.22%'],
                ['Software (Sistemas e Aplicações)', '462', '4.43%'],
        ]],
        ['center'=>  'Beta dos setores: 1.28 β']
    ],

    [ // slide2
        ['h1'=> 'Desempenho '],
        ['p1' => ['lines'=>1,'text'=>'Receita                                                  297.000,00 R$']],
        ['p2' => ['lines'=>1,'text'=>'Dívidas                                                  15.000,00 R$']],
        ['p3' => ['lines'=>1,'text'=>'Dinheiro em caixa                                        20.000,00 R$']],
        ['p4' => ['lines'=>1,'text'=>'Retorno Sobre o Capital Investido (ROIC)                 80,00 R$']],
        ['p5' => ['lines'=>1,'text'=>'EBIT                                                     84.800,00 R$']],
        ['p6' =>      ['lines'=>1,'text'=>'Margem EBIT                                              28.55%']],
        [ 'table'=> [
            [ ['Taxa Livre de Risco', '12.56%'], ['Prêmio de Risco de Capital Próprio', '7.21%'], ['Desconto de Iliquidez', '15%'] ],
            [ ['Risco País', '2.97%'], ['Custo do Capital', '19,28%'], ['Custo da Dívida', '15,88%'] ],
            [ ['Custo do Capital Próprio', '19,75%'], ['IPCA acumulado 12 meses', '11.30%'], ['Cotação diária do dólar', 'R$ 5,13'] ],
        ] ],
        [ 'h1'=> '1. Valor econômico com DCF' ],
        [ 'p1'=> ['text'=>'No DCF (Discounted Cash Flow, ou Fluxo de Caixa Descontado), o objetivo é identificar o valor de uma empresa considerando seu fluxo de caixa, crescimento e risco. Ou seja, estima o seu potencial econômico em gerar retorno para sócios e credores. Esse critério engloba fatores como plano de negócio, competência da gestão, tradição, marca, carteira de clientes e qualquer decisão que impacte na geração de resultados.', 'lines'=>6] ],
        [ 'p1'=>['text'=>'Conforme Aswath Damodaran, uma boa avaliação é o resultado de dois elementos: o desempenho histórico da empresa e a criação de uma narrativa sólida (através das projeções financeiras).', 'lines'=>3] ],
    ],

    [ // slide3
        ['centerBox'=> ['text'=>'Desempenho', 'text2'=> 'R$ 718.740,00', 'text2Font'=>20, 'bg'=>'FFd7f5fc', 'text2Color'=>'FF03c3ec', 'width'=>600 ]],
        ['p1' => ['lines'=>1,'text'=>'Receita                                                  297.000,00 R$']],
        ['p2' => ['lines'=>1,'text'=>'Dívidas                                                  15.000,00 R$']],
        ['p3' => ['lines'=>1,'text'=>'Dinheiro em caixa                                        20.000,00 R$']],
        ['p4' => ['lines'=>1,'text'=>'Retorno Sobre o Capital Investido (ROIC)                 80,00 R$']],
        ['p5' => ['lines'=>1,'text'=>'EBIT                                                     84.800,00 R$']],
        ['p6' =>      ['lines'=>1,'text'=>'Margem EBIT                                              28.55%']],
        [ 'table'=> [
            [ ['Taxa Livre de Risco', '12.56%'], ['Prêmio de Risco de Capital Próprio', '7.21%'], ['Desconto de Iliquidez', '15%'] ],
            [ ['Risco País', '2.97%'], ['Custo do Capital', '19,28%'], ['Custo da Dívida', '15,88%'] ],
            [ ['Custo do Capital Próprio', '19,75%'], ['IPCA acumulado 12 meses', '11.30%'], ['Cotação diária do dólar', 'R$ 5,13'] ],
        ] ],
        [ 'h1'=> '1. Valor econômico com DCF' ],
        [ 'p1'=> ['text'=>'No DCF (Discounted Cash Flow, ou Fluxo de Caixa Descontado), o objetivo é identificar o valor de uma empresa considerando seu fluxo de caixa, crescimento e risco. Ou seja, estima o seu potencial econômico em gerar retorno para sócios e credores. Esse critério engloba fatores como plano de negócio, competência da gestão, tradição, marca, carteira de clientes e qualquer decisão que impacte na geração de resultados.', 'lines'=>6] ],
        [ 'p1'=>['text'=>'Conforme Aswath Damodaran, uma boa avaliação é o resultado de dois elementos: o desempenho histórico da empresa e a criação de uma narrativa sólida (através das projeções financeiras).', 'lines'=>3] ],
    ],

];


createPpt($data);



function createPpt($slides, $title='slide'){
    $objPHPPowerPoint = new PhpPresentation();

    $i=0;
    foreach($slides as $slide){
        $currentSlide = $i==0 ? $objPHPPowerPoint->getActiveSlide() : $objPHPPowerPoint->createSlide();

        $currentYOffset = 0;        
        foreach($slide as $element){
            $tag = array_key_first($element);
            $value = $element[$tag];

            if($tag=='logo'){
                $currentYOffset += 10; 
                $height = 36;
                $shape = $currentSlide->createDrawingShape();
                $shape->setName('PHPPresentation logo')
                    ->setDescription('PHPPresentation logo')
                    ->setPath($value)
                    ->setHeight(36)
                    ->setOffsetX(400)
                    ->setOffsetY($currentYOffset);
                // $shape->getShadow()->setVisible(true)
                //                    ->setDirection(45)
                //                    ->setDistance(10);
            }

            if( in_array($tag ,['h1','h2','h3','h4']) ){
                $currentYOffset += 5; 
                $height = 40;
                $fontSize = $tag=='h1' ? 20 : 18; 
                $shape = $currentSlide->createRichTextShape()
                ->setHeight($height)
                ->setWidth(600)
                ->setOffsetX(10)
                ->setOffsetY($currentYOffset);
                $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
                $textRun = $shape->createTextRun($value);
                $textRun->getFont()->setBold(true)
                ->setSize($fontSize)
                ->setColor( new Color( 'FF47008a' ) );

            }

            if( in_array($tag ,['p1','p2','p2_line','p3','p4','p5','p6', 'center']) ){
                $currentYOffset += 5; 
                $height = 40;
                if(is_array($value)){
                    $height = 25*$value['lines'];
                    $value = $value['text'];
                }
                
                $fontSize = 14; 
                $shape = $currentSlide->createRichTextShape()
                ->setHeight($height)
                ->setWidth(800)
                ->setOffsetX(10)
                ->setOffsetY($currentYOffset);
                if($tag=='center'){
                    $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
                }else{
                    $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
                }
                $textRun = $shape->createTextRun($value);
                $textRun->getFont()
                ->setSize($fontSize)
                ->setColor( new Color( 'FF000000' ) );

                if($tag=='p2_line'){
                    echo 'p2_line';
                    // $shape->getBorder()
                    // ->setLineStyle(Border::LINE_SINGLE)
                    // ->setDashStyle(Border::DASH_DASH)
                    // ->setLineWidth(4)
                    // ->getColor()->setARGB('FFC00000');

                }

            }

            if($tag=='centerBox'){
                $currentYOffset += 5; 
                $height = 40;
                $width = $value['width'];
                $text = $value['text'];
                // if(is_array($value)){
                //     $height = 25*$value['lines'];
                //     $value = $value['text'];
                // }
                
                $fontSize = 14; 
                $shape = $currentSlide->createRichTextShape()
                ->setHeight($height)
                ->setWidth($width)
                ->setOffsetX(150)
                ->setOffsetY($currentYOffset);
                $shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
                $textRun = $shape->createTextRun($text);
                $textRun->getFont()
                ->setSize($fontSize)
                ->setColor( new Color( 'FF000000' ) );
                $bgColor = new Color($value['bg']);
                $shape->getFill()->setFillType(Fill::FILL_SOLID)
                    ->setStartColor($bgColor)
                    ->setEndColor($bgColor);
            }

            if($tag == 'table'){
                $currentYOffset += 15;
                $height = sizeof($value)*40;
                if( is_array($value[0][0]) ){
                    $height = sizeof($value)*50;
                }
                $columns = $value[0];
                $shape = $currentSlide->createTableShape( sizeof($columns) );
                $shape->setOffsetY($currentYOffset)
                ->setOffsetX(10)
                ->setWidth(800);

                createTable($shape, $value);
            }

            $currentYOffset += $height;
        }

        $i++;
    }
    
    
    $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
    $oWriterPPTX->save(__DIR__ . "/$title.pptx");
    
}

function createTable($shape, $table){
    $j=0;

    $isTwoDimenional = is_array($table[0][0]);
    foreach($table as $rowData){
        $row = $shape->createRow();
        foreach($rowData as $cell){
            $cellValue = $isTwoDimenional ? $cell[0] : $cell;
            $secondLine = $isTwoDimenional ? $cell[1] : '';
            $colorHex = $j==0 && !$isTwoDimenional ? 'FF47008a' : 'FF000000';

            $cell = $row->nextCell();
            $textRun = $cell->createTextRun($cellValue);
            $textRun->getFont()->setBold(true);
            $textRun->getFont()->setSize(12);
            $textRun->getFont()->setColor(new Color( $colorHex ));
            if($isTwoDimenional){
                $cell->createBreak();
                $textRun = $cell->createTextRun($secondLine);
                $textRun->getFont()->setBold(true);
                $textRun->getFont()->setSize(14);
                $textRun->getFont()->setColor(new Color('FF47008a'));
            }

            $cell->getActiveParagraph()->getAlignment()
            ->setMarginLeft(10)->setMarginBottom(5);
        }
        $j++;
    }
}