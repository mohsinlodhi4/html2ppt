<?php

require __DIR__.'/vendor/autoload.php';

\PhpOffice\PhpPresentation\Autoloader::register();
\PhpOffice\Common\Autoloader::register();

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;

// echo "<pre>";
// print_r( get_class_methods(PhpPresentation::class) );
createImage();
$slideImages = cropImage(__DIR__.'/result.jpg');
createPpt($slideImages);



function createPpt($slides, $title='slide'){
    $objPHPPowerPoint = new PhpPresentation();

    $i=0;
    foreach($slides as $slide){
        $currentSlide = $i==0 ? $objPHPPowerPoint->getActiveSlide() : $objPHPPowerPoint->createSlide();
        
        // Create a shape (drawing)
        $shape = $currentSlide->createDrawingShape();
        $shape->setName('PHPPresentation logo')
            ->setDescription('PHPPresentation logo')
            ->setPath($slide)
            ->setHeight(800)
            ->setWidth(800)
            ->setOffsetX(50)
            ->setOffsetY(50);
        $shape->getShadow()->setVisible(true)
                        ->setDirection(45)
                        ->setDistance(10);

        // if($i!=0) $objPHPPowerPoint->addSlide($currentSlide);

        $i++;
    }
    
    
    $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
    $oWriterPPTX->save(__DIR__ . "/sample.pptx");
    
}

function createImage($url=''){
    $url = $url=='' ? "https://valutech.io/uploads/modelo.htm" : $url;
    try
    {
        // create the API client instance
        $client = new \Pdfcrowd\HtmlToImageClient("demo", "ce544b6ea52a5621fb9d55f8b542d14d");

        // configure the conversion
        $client->setOutputFormat("jpg");
        $client->setCustomJavascript("libPdfcrowd.removeZIndexHigherThan({zlimit: 90});");

        // run the conversion and write the result to a file
        $client->convertUrlToFile($url, "result.jpg");
    }
    catch(\Pdfcrowd\Error $why)
    {
        // report the error
        error_log("Pdfcrowd Error: {$why}\n");

        // rethrow or handle the exception
        throw $why;
    }

}


function cropImage($path, $prefix=''){
    $prefix = $prefix=='' ? 'pptImages' : $prefix;
    
    $im = imagecreatefromjpeg($path);
    $pageHeight = 795;
    $imageHeight = imagesy($im);
    
    $images = [];
    // find the size of image
    $size = min(imagesx($im), imagesy($im));
    if(! is_dir(__DIR__."/$prefix")){
        mkdir(__DIR__."/$prefix");
    }

    $i=1;
    $y=0;
    while($y<$imageHeight){
        if($i==4){
            $pageHeight = 600;
        }
        // elseif($i==5){
        //     $pageHeight = 650;
        // }
        // elseif($i==6){
        //     $pageHeight = 662;
        // }
        elseif($i==8){
            $pageHeight = 805;

        }
        elseif($i==9){
            break;
        }
        else{
            $pageHeight = 795;
        }
        $isLast = $y+$pageHeight>$imageHeight ? true : false;
        // $height = $isLast ? 570 : $pageHeight;
        $height = $pageHeight;
        // Set the crop image size 
        $im2 = imagecrop($im, ['x' => 0, 'y' => $y, 'width' => 1010, 'height' => $height]);
        if ($im2 !== FALSE) {
            // header("Content-type: image/png");
            $name =__DIR__."/$prefix/".generateRandomString(10).".png";
            $images[] =  $name;
            imagepng($im2, $name);
            imagedestroy($im2);
        }
        imagedestroy($im);

        $y+=$pageHeight;
        $i++;
    }

    return $images;
      
}

function generateRandomString($length = 10) {
    $characters = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ';
    $charactersLength = strlen($characters);
    $randomString = '';
    for ($i = 0; $i < $length; $i++) {
        $randomString .= $characters[rand(0, $charactersLength - 1)];
    }
    return $randomString;
}