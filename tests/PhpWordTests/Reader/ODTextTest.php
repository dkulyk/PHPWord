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
use PhpOffice\PhpWord\Element\ListItem;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\Text;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Style;
use PhpOffice\PhpWord\Style\Numbering;
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

    public function testListFormsOfLibreOffice(): void
    {
        // A level with its indent as space before and label width, and one with the whole text
        // of its number; a list header, an item of two paragraphs, and a list of no known style
        $styles = '<text:list-style style:name="L1">'
            . '<text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix=".">'
            . '<style:list-level-properties text:space-before="0.5in" text:min-label-width="0.25in"/>'
            . '<style:text-properties style:font-name="Arial"/></text:list-level-style-number>'
            . '<text:list-level-style-number text:level="2" style:num-format="a" style:num-suffix=")" text:display-levels="2" loext:num-list-format="%1%-%2%)"/>'
            . '</text:list-style>';
        $body = '<text:list text:style-name="L1">'
            . '<text:list-header><text:p>Header</text:p></text:list-header>'
            . '<text:list-item><text:p>One</text:p><text:p>More</text:p>'
            . '<text:list><text:list-item><text:p>Two</text:p></text:list-item></text:list></text:list-item>'
            . '</text:list><text:list text:style-name="Nowhere"><text:list-item><text:p>Three</text:p></text:list-item></text:list>';
        $file = (string) tempnam(sys_get_temp_dir(), 'PhpWord');
        $zip = new ZipArchive();
        $zip->open($file, ZipArchive::OVERWRITE);
        $zip->addFromString('mimetype', 'application/vnd.oasis.opendocument.text');
        $namespaces = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';
        $zip->addFromString('styles.xml', '<office:document-styles ' . $namespaces . '><office:styles/></office:document-styles>');
        $zip->addFromString('content.xml', '<office:document-content ' . $namespaces . '><office:automatic-styles>' . $styles . '</office:automatic-styles><office:body><office:text>' . $body . '</office:text></office:body></office:document-content>');
        $zip->close();

        $phpWord = IOFactory::load($file, 'ODText');
        unlink($file);

        $read = [];
        foreach ($phpWord->getSections()[0]->getElements() as $element) {
            $read[] = $element instanceof ListItem
                ? [$element->getText(), $element->getDepth(), $element->getStyle()->getNumStyle()]
                : [$element instanceof Text ? $element->getText() : get_class($element)];
        }
        self::assertEquals([['Header'], ['One', 0, 'L1'], ['More'], ['Two', 1, 'L1'], ['Three', 0, 'PHPWordListType3']], $read);

        $numbering = Style::getStyle('L1');
        self::assertInstanceOf(Numbering::class, $numbering);
        $levels = [];
        foreach ($numbering->getLevels() as $level) {
            $levels[] = [$level->getText(), $level->getLeft(), $level->getHanging(), $level->getTabPos(), $level->getFont()];
        }
        self::assertEquals([['%1.', 1080, 360, 1080, 'Arial'], ['%1-%2)', null, null, null, null]], $levels);
    }

    public function testNumberedListSurvivesTheRoundTrip(): void
    {
        $phpWord = new PhpWord();
        $phpWord->addNumberingStyle('Num', [
            'type' => 'multilevel',
            'levels' => [
                ['format' => 'decimal', 'text' => '%1.', 'start' => 3],
                ['format' => 'lowerLetter', 'text' => '(%2)'],
                ['format' => 'upperRoman', 'text' => '%1.%2.%3'],
                ['format' => 'bullet', 'text' => '•'],
                ['format' => 'none', 'text' => 'Note'],
            ],
        ]);
        $section = $phpWord->addSection();
        foreach ([0, 1, 2, 1, 0] as $i => $depth) {
            $section->addListItem('Item ' . $i, $depth, null, 'Num');
        }
        $file = (string) tempnam(sys_get_temp_dir(), 'PhpWord');
        IOFactory::createWriter($phpWord, 'ODText')->save($file);
        Style::resetStyles();

        $phpWordRead = IOFactory::load($file, 'ODText');
        unlink($file);

        $items = [];
        foreach ($phpWordRead->getSections()[0]->getElements() as $element) {
            self::assertInstanceOf(ListItem::class, $element);
            $items[] = [$element->getText(), $element->getDepth(), $element->getStyle()->getNumStyle()];
        }
        self::assertEquals([
            ['Item 0', 0, 'Num'],
            ['Item 1', 1, 'Num'],
            ['Item 2', 2, 'Num'],
            ['Item 3', 1, 'Num'],
            ['Item 4', 0, 'Num'],
        ], $items);

        $numbering = Style::getStyle('Num');
        self::assertInstanceOf(Numbering::class, $numbering);
        $levels = [];
        foreach ($numbering->getLevels() as $level) {
            $levels[] = [$level->getFormat(), $level->getText(), $level->getStart()];
        }
        self::assertEquals([
            ['decimal', '%1.', 3],
            ['lowerLetter', '(%2)', 1],
            ['upperRoman', '%1.%2.%3', 1],
            ['bullet', '•', 1],
            ['none', 'Note', 1],
        ], $levels);
    }
}
