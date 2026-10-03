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

namespace PhpOffice\PhpWordTests\Reader\Word2007;

use PhpOffice\PhpWord\Element\Image;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Settings;
use PHPUnit\Framework\TestCase;
use ZipArchive;

class ImageAltTextTest extends TestCase
{
    /** @var string */
    private $filename = '';

    protected function tearDown(): void
    {
        if ($this->filename !== '') {
            unlink($this->filename);
            $this->filename = '';
        }
    }

    /**
     * Save a document with one image, and give back the relationship id of the image.
     */
    private function saveImage(bool $decorative = false): string
    {
        $phpWord = new PhpWord();
        $phpWord->addSection()->addImage(__DIR__ . '/../../_files/images/earth.jpg', null, false, 'Earth', 'The Earth seen from space')->setDecorative($decorative);
        $this->filename = (string) tempnam(Settings::getTempDir(), 'PhpWord');
        IOFactory::createWriter($phpWord, 'Word2007')->save($this->filename);

        $zip = new ZipArchive();
        $zip->open($this->filename);
        self::assertSame(1, preg_match('/<a:blip r:embed="([^"]+)"/', (string) $zip->getFromName('word/document.xml'), $matches));
        $zip->close();

        return $matches[1] ?? '';
    }

    private function readImage(): Image
    {
        $textRun = IOFactory::load($this->filename)->getSection(0)->getElement(0);
        self::assertInstanceOf(TextRun::class, $textRun);
        $image = $textRun->getElement(0);
        self::assertInstanceOf(Image::class, $image);

        return $image;
    }

    /**
     * Replace the w:drawing of the saved image with the given markup.
     */
    private function replaceDrawing(string $markup): void
    {
        $zip = new ZipArchive();
        $zip->open($this->filename);
        $document = (string) $zip->getFromName('word/document.xml');
        $zip->addFromString('word/document.xml', (string) preg_replace('#<w:drawing>.*</w:drawing>#s', $markup, $document));
        $zip->close();
    }

    public function testReadWrittenAltText(): void
    {
        $this->saveImage();

        self::assertSame('The Earth seen from space', $this->readImage()->getAltText());
    }

    public function testReadVmlAltText(): void
    {
        // As PHPWord 1.4 writes it
        $rId = $this->saveImage();
        $this->replaceDrawing('<w:pict><v:shape type="#_x0000_t75" stroked="f" alt="The Earth seen from space" style="width:75pt;height:75pt;"><v:imagedata r:id="' . $rId . '" o:title=""/></v:shape></w:pict>');

        self::assertSame('The Earth seen from space', $this->readImage()->getAltText());
    }

    public static function providerDrawing(): array
    {
        $graphic = '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            . '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="earth.jpg"%s/><pic:cNvPicPr/></pic:nvPicPr>'
            . '<pic:blipFill><a:blip r:embed="%s"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>';
        $inline = '<wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Picture 1"%s/>' . $graphic . '</wp:inline>';
        $anchor = '<wp:anchor behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" relativeHeight="1" simplePos="0" distT="0" distB="0" distL="0" distR="0">'
            . '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            . '<wp:extent cx="952500" cy="952500"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Picture 1"%s/>' . $graphic . '</wp:anchor>';
        $descr = ' descr="The Earth seen from space"';

        return [
            // The alternative text in wp:docPr only
            'inline, docPr' => [$inline, $descr, ''],
            'anchor, docPr' => [$anchor, $descr, ''],
            // In both, as LibreOffice writes it
            'inline, both' => [$inline, $descr, $descr],
            // Read before this change
            'inline, pic:cNvPr only' => [$inline, '', $descr],
        ];
    }

    /**
     * @dataProvider providerDrawing
     */
    public function testReadDrawingAltText(string $frame, string $docPrDescr, string $cNvPrDescr): void
    {
        $rId = $this->saveImage();
        $this->replaceDrawing('<w:drawing>' . sprintf($frame, $docPrDescr, $cNvPrDescr, $rId) . '</w:drawing>');

        $image = $this->readImage();
        self::assertSame('The Earth seen from space', $image->getAltText());
        self::assertSame('earth.jpg', $image->getName());
    }

    public function testReadWrittenDecorative(): void
    {
        $this->saveImage(true);

        self::assertTrue($this->readImage()->isDecorative());
    }

    public static function providerDecorative(): array
    {
        $extLst = '<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">'
            . '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="%s"/></a:ext></a:extLst>';
        // The frames of providerDrawing, with the content of wp:docPr in place of its descr
        $drawing = self::providerDrawing();
        [$inline, $anchor] = str_replace('name="Picture 1"%s/>', 'name="Picture 1">%s</wp:docPr>', [$drawing['inline, docPr'][0], $drawing['anchor, docPr'][0]]);

        return [
            // As Word writes it
            'inline' => [$inline, sprintf($extLst, '1'), true],
            'anchor' => [$anchor, sprintf($extLst, '1'), true],
            'true' => [$inline, sprintf($extLst, 'true'), true],
            '0' => [$inline, sprintf($extLst, '0'), false],
            'false' => [$inline, sprintf($extLst, 'false'), false],
            'xsd:boolean whitespace' => [$inline, sprintf($extLst, ' true '), true],
            // Word puts its a16:creationId extension first
            'after another extension' => [$inline, str_replace('<a:ext uri="{C183', '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{00000000-0000-0000-0000-000000000001}"/></a:ext><a:ext uri="{C183', sprintf($extLst, '1')), true],
            'no extension' => [$inline, '', false],
        ];
    }

    /**
     * @dataProvider providerDecorative
     */
    public function testReadDecorative(string $frame, string $extLst, bool $decorative): void
    {
        $rId = $this->saveImage();
        $this->replaceDrawing('<w:drawing>' . sprintf($frame, $extLst, '', $rId) . '</w:drawing>');

        self::assertSame($decorative, $this->readImage()->isDecorative());
    }
}
