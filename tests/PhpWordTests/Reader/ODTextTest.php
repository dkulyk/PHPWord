<?php

/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @see         https://github.com/PHPOffice/PHPWord
 *
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWordTests\Reader;

use PhpOffice\Math\Element;
use PhpOffice\PhpWord\Element\Formula;
use PhpOffice\PhpWord\Element\Image;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\Element\TextBreak;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Reader\ODText;
use ZipArchive;

/**
 * Test class for PhpOffice\PhpWord\Reader\ODText.
 *
 * @coversDefaultClass \PhpOffice\PhpWord\Reader\ODText
 *
 * @runTestsInSeparateProcesses
 */
class ODTextTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Load.
     */
    public function testLoad(): void
    {
        $phpWord = IOFactory::load(dirname(__DIR__, 1) . '/_files/documents/reader.odt', 'ODText');
        self::assertInstanceOf(PhpWord::class, $phpWord);
    }

    public function testLoadFormula(): void
    {
        $phpWord = IOFactory::load(dirname(__DIR__, 1) . '/_files/documents/reader-formula.odt', 'ODText');

        self::assertInstanceOf(PhpWord::class, $phpWord);

        $sections = $phpWord->getSections();
        self::assertCount(1, $sections);

        $section = $sections[0];
        self::assertInstanceOf(Section::class, $section);

        $elements = $section->getElements();
        self::assertCount(1, $elements);

        $element = $elements[0];
        self::assertInstanceOf(Formula::class, $element);

        $elements = $element->getMath()->getElements();
        self::assertCount(1, $elements);

        self::assertInstanceOf(Element\Semantics::class, $elements[0]);
    }

    public function testImagesInTheFormsOfLibreOffice(): void
    {
        $image = '<draw:image xlink:href="Pictures/%s" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"><text:p/></draw:image>';
        $body = '<text:p><draw:frame draw:name="Earth" text:anchor-type="as-char" svg:width="1.25in" svg:height=".5in">'
            . sprintf($image, 'earth.svg') . sprintf($image, 'earth.jpg')
            . '<svg:title>Earth</svg:title><svg:desc>The Earth from space</svg:desc></draw:frame></text:p>'
            . '<text:p>Before <draw:frame text:anchor-type="as-char" svg:width="0.5cm" svg:height="5mm">' . sprintf($image, 'mars.jpg') . '<svg:desc>Mars</svg:desc></draw:frame><text:s/>after</text:p>'
            . '<text:p><draw:frame text:anchor-type="paragraph" svg:width="4cm"><draw:text-box><text:p>'
            . '<draw:frame svg:width="3cm" svg:height="3cm">' . sprintf($image, 'mars.jpg') . '</draw:frame>Illustration <text:sequence text:name="Illustration">1</text:sequence>: Mars</text:p></draw:text-box></draw:frame></text:p>'
            . '<text:p><draw:a xlink:href="https://example.org/"><draw:frame svg:width="1cm" svg:height="1cm">' . sprintf($image, 'my%20pic.jpg') . '<svg:title>Only a title</svg:title></draw:frame></draw:a>'
            . '<text:a xlink:href="https://example.org/">link</text:a><text:line-break/>next<text:note><text:note-citation>1</text:note-citation><text:note-body><text:p>Note</text:p></text:note-body></text:note></text:p>'
            . '<text:p><text:span>Formula <draw:frame><draw:object xlink:href="./Object 1"/>' . sprintf($image, 'mars.jpg') . '</draw:frame></text:span></text:p>'
            . '<text:p>Outside<draw:frame svg:width="in">' . str_replace('Pictures/%s', '../mars.jpg', $image) . str_replace('Pictures/%s', 'file:///etc/hosts', $image)
            . str_replace('Pictures/%s', '/Pictures/mars.jpg', $image) . '</draw:frame></text:p>';
        $file = (string) tempnam(sys_get_temp_dir(), 'PhpWord');
        $zip = new ZipArchive();
        $zip->open($file, ZipArchive::OVERWRITE);
        $zip->addFromString('mimetype', 'application/vnd.oasis.opendocument.text');
        $zip->addFromString('content.xml', '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text>' . $body . '</office:text></office:body></office:document-content>');
        $zip->addFromString('Pictures/earth.svg', '<svg xmlns="http://www.w3.org/2000/svg"/>');
        $zip->addFile(dirname(__DIR__) . '/_files/images/earth.jpg', 'Pictures/earth.jpg');
        $zip->addFile(dirname(__DIR__) . '/_files/images/mars.jpg', 'Pictures/mars.jpg');
        $zip->addFile(dirname(__DIR__) . '/_files/images/mars.jpg', 'Pictures/my pic.jpg');
        $zip->close();

        $phpWord = IOFactory::load($file, 'ODText');
        unlink($file);

        self::assertEquals([
            [['Pictures/earth.jpg', 120.0, 48.0, 'px', 'Earth', 'Earth - The Earth from space']],
            ['Before ', ['Pictures/mars.jpg', 18.9, 18.9, 'px', null, 'Mars'], ' ', 'after'],
            [['Pictures/mars.jpg', 113.39, 113.39, 'px', null, null], 'Illustration ', '1', ': Mars'],
            [['Pictures/my pic.jpg', 37.8, 37.8, 'px', null, 'Only a title'], 'link', 'BR', 'next'],
            ['Formula '],
            ['Outside'],
        ], $this->readRuns($phpWord));
    }

    public function testImageSurvivesTheRoundTrip(): void
    {
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        $section->addImage(dirname(__DIR__) . '/_files/images/earth.jpg', ['width' => 120, 'height' => 90, 'unit' => 'px'], false, null, 'The Earth from space');
        $textRun = $section->addTextRun();
        $textRun->addText('Before ');
        $textRun->addImage(dirname(__DIR__) . '/_files/images/mars.jpg', ['width' => 30, 'height' => 30, 'unit' => 'px']);
        $file = (string) tempnam(sys_get_temp_dir(), 'PhpWord');
        IOFactory::createWriter($phpWord, 'ODText')->save($file);

        $read = IOFactory::load($file, 'ODText');
        $withoutImages = (new ODText())->setImageLoading(false)->load($file);
        unlink($file);

        self::assertEquals([
            [['Pictures/section_image1.jpg', 120.0, 90.0, 'px', null, 'The Earth from space']],
            ['Before', ' ', ['Pictures/section_image2.jpg', 30.0, 30.0, 'px', null, null]],
        ], $this->readRuns($read, false));
        self::assertEquals([[], ['Before', ' ']], $this->readRuns($withoutImages, false));
    }

    /**
     * The runs of the first section: a text as its string, an image as its part, size, name and alternative text.
     *
     * @return array<int, array<int, mixed>>
     */
    private function readRuns(PhpWord $phpWord, bool $withName = true): array
    {
        $runs = [];
        foreach ($phpWord->getSections()[0]->getElements() as $run) {
            self::assertInstanceOf(TextRun::class, $run);
            $read = [];
            foreach ($run->getElements() as $element) {
                if ($element instanceof Image) {
                    $style = $element->getStyle();
                    $read[] = [
                        (string) substr($element->getSource(), (int) strpos($element->getSource(), '#') + 1),
                        round((float) $style->getWidth(), 2),
                        round((float) $style->getHeight(), 2),
                        $style->getUnit(),
                        $withName ? $element->getName() : null,
                        $element->getAltText(),
                    ];
                } elseif ($element instanceof TextBreak) {
                    $read[] = 'BR';
                } elseif ($element instanceof Text) {
                    $read[] = $element->getText();
                }
            }
            $runs[] = $read;
        }

        return $runs;
    }
}
