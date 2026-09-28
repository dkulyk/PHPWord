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

namespace PhpOffice\PhpWordTests\Writer\ODText\Style;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\Jc;
use PhpOffice\PhpWordTests\TestHelperDOCX;

class NumberingTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Executed after each method of the class.
     */
    protected function tearDown(): void
    {
        TestHelperDOCX::clear();
    }

    public function testAddListItemRun(): void
    {
        $expected = 'MyOwnNumberingStyle';

        $phpWord = new PhpWord();
        $phpWord->addNumberingStyle($expected, [
            'type' => 'multilevel',
            'levels' => [
                [
                    'start' => 1,
                    'format' => 'decimal',
                    'restart' => 1,
                    'suffix' => 'space',
                    'text' => '%1.',
                    'alignment' => Jc::START,
                ],
            ],
        ]);
        $phpWord->addSection()
            ->addListItemRun(0, $expected)
            ->addText('List item run 1');

        $doc = TestHelperDOCX::getDocument($phpWord, 'ODText');
        $doc->setDefaultFile('styles.xml');

        $xPath = '/office:document-styles/office:styles';
        self::assertTrue($doc->elementExists($xPath));
        self::assertTrue($doc->elementExists($xPath . '/text:list-style'));
        self::assertTrue($doc->hasElementAttribute($xPath . '/text:list-style', 'style:name'));
        self::assertEquals($expected, $doc->getElementAttribute($xPath . '/text:list-style', 'style:name'));
        self::assertTrue($doc->elementExists($xPath . '/text:list-style/text:list-level-style-number'));
        self::assertTrue($doc->elementExists($xPath . '/text:list-style/text:list-level-style-number/style:list-level-properties'));
        self::assertTrue($doc->elementExists($xPath . '/text:list-style/text:list-level-style-number/style:list-level-properties/style:list-level-label-alignment'));
        self::assertTrue($doc->elementExists($xPath . '/text:list-style/text:list-level-style-number/style:text-properties'));
    }

    public function testNumberFormat(): void
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
                ['format' => 'ordinal', 'text' => '%6.'],
            ],
        ]);
        $phpWord->addSection()->addListItem('Item', 0, null, 'Num');

        $doc = TestHelperDOCX::getDocument($phpWord, 'ODText');
        $doc->setDefaultFile('styles.xml');

        $attributes = ['style:num-format', 'style:num-prefix', 'style:num-suffix', 'text:display-levels', 'style:num-letter-sync', 'text:start-value', 'text:bullet-char'];
        $levels = [];
        for ($level = 1; $level <= 6; ++$level) {
            foreach (['number', 'bullet'] as $kind) {
                $xPath = '/office:document-styles/office:styles/text:list-style/text:list-level-style-' . $kind . '[@text:level="' . $level . '"]';
                if ($doc->elementExists($xPath)) {
                    foreach ($attributes as $attribute) {
                        if ($doc->hasElementAttribute($xPath, $attribute)) {
                            $levels[$level][$attribute] = $doc->getElementAttribute($xPath, $attribute);
                        }
                    }
                }
            }
        }

        // Letters repeat as Word's do, and a format OpenDocument does not have is written in arabic numbers
        self::assertEquals([
            1 => ['style:num-format' => '1', 'style:num-suffix' => '.', 'text:start-value' => '3'],
            2 => ['style:num-format' => 'a', 'style:num-prefix' => '(', 'style:num-suffix' => ')', 'style:num-letter-sync' => 'true', 'text:start-value' => '1'],
            3 => ['style:num-format' => 'I', 'text:display-levels' => '3', 'text:start-value' => '1'],
            4 => ['text:bullet-char' => '•'],
            5 => ['style:num-format' => '', 'style:num-prefix' => 'Note', 'text:start-value' => '1'],
            6 => ['style:num-format' => '1', 'style:num-suffix' => '.', 'text:start-value' => '1'],
        ], $levels);
    }
}
