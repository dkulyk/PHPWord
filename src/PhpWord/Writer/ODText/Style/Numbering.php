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

namespace PhpOffice\PhpWord\Writer\ODText\Style;

use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Style\Numbering as StyleNumbering;
use PhpOffice\PhpWord\Style\NumberingLevel;

/**
 * Numbering style writer.
 */
class Numbering extends AbstractStyle
{
    /**
     * The `style:num-format` of each number format of Word that OpenDocument has, and whether its
     * letters repeat (`aa`, `bb`) as Word's do; any other format is written in arabic numbers.
     *
     * @var array<string, array{0: string, 1: bool}>
     */
    public const NUM_FORMATS = [
        'decimal' => ['1', false],
        'lowerLetter' => ['a', true],
        'upperLetter' => ['A', true],
        'lowerRoman' => ['i', false],
        'upperRoman' => ['I', false],
        'none' => ['', false],
    ];

    /**
     * Write style.
     */
    public function write(): void
    {
        /** @var StyleNumbering $style Type hint */
        $style = $this->getStyle();
        if (!$style instanceof StyleNumbering) {
            return;
        }
        $xmlWriter = $this->getXmlWriter();

        $xmlWriter->startElement('text:list-style');
        $xmlWriter->writeAttribute('style:name', $style->getStyleName());

        foreach ($style->getLevels() as $styleLevel) {
            $numLevel = $styleLevel->getLevel() + 1;

            // In Twips
            $tabPos = $styleLevel->getTabPos();
            // In Inches
            $tabPos /= Converter::INCH_TO_TWIP;
            // In Centimeters
            $tabPos *= Converter::INCH_TO_CM;

            // In Twips
            $hanging = $styleLevel->getHanging();
            // In Inches
            $hanging /= Converter::INCH_TO_TWIP;
            // In Centimeters
            $hanging *= Converter::INCH_TO_CM;

            $isBullet = $styleLevel->getFormat() === 'bullet';
            $xmlWriter->startElement($isBullet ? 'text:list-level-style-bullet' : 'text:list-level-style-number');
            $xmlWriter->writeAttribute('text:level', $numLevel);
            $xmlWriter->writeAttribute('text:style-name', $style->getStyleName() . '_' . $numLevel);
            if ($isBullet) {
                $xmlWriter->writeAttribute('text:bullet-char', $styleLevel->getText());
            } else {
                $this->writeNumber($styleLevel);
            }

            $xmlWriter->startElement('style:list-level-properties');
            $xmlWriter->writeAttribute('text:list-level-position-and-space-mode', 'label-alignment');

            $xmlWriter->startElement('style:list-level-label-alignment');
            $xmlWriter->writeAttribute('text:label-followed-by', 'listtab');
            $xmlWriter->writeAttribute('text:list-tab-stop-position', number_format($tabPos, 2, '.', '') . 'cm');
            $xmlWriter->writeAttribute('fo:text-indent', '-' . number_format($hanging, 2, '.', '') . 'cm');
            $xmlWriter->writeAttribute('fo:margin-left', number_format($tabPos, 2, '.', '') . 'cm');

            $xmlWriter->endElement(); // style:list-level-label-alignment
            $xmlWriter->endElement(); // style:list-level-properties

            $xmlWriter->startElement('style:text-properties');
            $xmlWriter->writeAttribute('style:font-name', $styleLevel->getFont());
            $xmlWriter->endElement(); // style:text-properties

            $xmlWriter->endElement(); // text:list-level-style-bullet|number
        }

        $xmlWriter->endElement(); // text:list-style
    }

    /**
     * Write the number of a level: Word's text such as `%1.%2)` is the numbers of the levels it
     * shows, with the text before the first and after the last as prefix and suffix.
     */
    private function writeNumber(NumberingLevel $styleLevel): void
    {
        $xmlWriter = $this->getXmlWriter();
        $text = (string) $styleLevel->getText();
        [$numFormat, $letterSync] = self::NUM_FORMATS[(string) $styleLevel->getFormat()] ?? self::NUM_FORMATS['decimal'];
        if (preg_match_all('/%\d/', $text, $matches, PREG_OFFSET_CAPTURE) > 0) {
            $first = $matches[0][0][1];
            $last = $matches[0][count($matches[0]) - 1][1] + 2;
            $xmlWriter->writeAttributeIf($first > 0, 'style:num-prefix', substr($text, 0, $first));
            $xmlWriter->writeAttributeIf($last < strlen($text), 'style:num-suffix', substr($text, $last));
            $xmlWriter->writeAttributeIf(count($matches[0]) > 1, 'text:display-levels', (string) count($matches[0]));
        } else {
            // A text with no number in it shows no number
            $numFormat = '';
            $letterSync = false;
            $xmlWriter->writeAttributeIf($text !== '', 'style:num-prefix', $text);
        }
        $xmlWriter->writeAttribute('style:num-format', $numFormat);
        $xmlWriter->writeAttributeIf($letterSync, 'style:num-letter-sync', 'true');
        // OpenDocument counts from 1 at the least
        $xmlWriter->writeAttributeIf($styleLevel->getStart() > 0, 'text:start-value', (string) $styleLevel->getStart());
    }
}
