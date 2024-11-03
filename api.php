<?php 


use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Shape\RichText;
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: POST, GET, OPTIONS, PUT, DELETE");
header("Access-Control-Allow-Headers: Content-Type");
require 'vendor/autoload.php';

$method = $_SERVER['REQUEST_METHOD'];

switch ($method) {
    case 'POST':
        $data = json_decode(file_get_contents('php://input'), true);
        
        if ($data && isset($data['content']) && !empty($data['content'])) {
            
            $pptFile = 'presentation.pptx';
            if (file_exists($pptFile)) {
                $ppt = IOFactory::load($pptFile);
            } else {
                $ppt = new PhpPresentation();
            }

            $slide = $ppt->createSlide();

            $shape = $slide->createRichTextShape();
            $shape->createTextRun($data['content']);

            $oWriterPPTX = IOFactory::createWriter($ppt, 'PowerPoint2007');
            $oWriterPPTX->save($pptFile);
            echo json_encode(['message' => 'New slide added successfully!']);
        } else {
            http_response_code(400);
            echo json_encode(['error' => 'Content is missing or empty']);
        }
        break;
    case 'GET':
        if (file_exists('presentation.pptx')) {
            $ppt = IOFactory::load('presentation.pptx');
            $slidesContent = [];
            foreach ($ppt->getAllSlides() as $index => $slide) {
                $content = '';
                foreach ($slide->getShapeCollection() as $shape) {
                    if ($shape instanceof RichText) {
                        $content .= $shape->getPlainText();
                    }
                }
                $slidesContent[] = ['slide' => $index + 1, 'content' => $content];
            }
            echo json_encode($slidesContent);
        } else {
            http_response_code(404); // File not found
            echo json_encode(['error' => 'Presentation not found']);
        }
        break;
        case 'PUT':
            $data = json_decode(file_get_contents('php://input'), true);
            if (file_exists('presentation.pptx') && isset($data['slide']) && isset($data['content'])) {
                $ppt = IOFactory::load('presentation.pptx');
                $slideIndex = $data['slide'] - 1; 
                if (isset($ppt->getAllSlides()[$slideIndex])) {
                    $slide = $ppt->getSlide($slideIndex);
        
                    foreach ($slide->getShapeCollection() as $shape) {
                        if ($shape instanceof PhpOffice\PhpPresentation\Shape\RichText) {
                            $shape->setParagraphs([]);
                            $shape->createTextRun($data['content']);
                            break;
                        }
                    }
                    $oWriterPPTX = IOFactory::createWriter($ppt, 'PowerPoint2007');
                    $oWriterPPTX->save('presentation.pptx');
                            echo json_encode(['message' => 'Slide updated successfully!']);
                } else {
                    http_response_code(404); // Slide not found
                    echo json_encode(['error' => 'Slide not found']);
                }
            } else {
                http_response_code(400); // Bad request
                echo json_encode(['error' => 'Invalid input or missing data']);
            }
            break;
    case 'DELETE':
        $data = json_decode(file_get_contents('php://input'), true);
        if (file_exists('presentation.pptx') && isset($data['slide'])) {
            $ppt = IOFactory::load('presentation.pptx');
            $slideIndex = $data['slide'] - 1; 
            if (isset($ppt->getAllSlides()[$slideIndex])) {
                $ppt->removeSlideByIndex($slideIndex);
                $oWriterPPTX = IOFactory::createWriter($ppt, 'PowerPoint2007');
                $oWriterPPTX->save('presentation.pptx');
                echo json_encode(['message' => 'Slide deleted successfully!']);
            } else {
                http_response_code(404); // Slide not found
                echo json_encode(['error' => 'Slide not found']);
            }
        } else {
            http_response_code(400); // Bad request
            echo json_encode(['error' => 'Invalid input or missing data']);
        }
        break;
}
?>
