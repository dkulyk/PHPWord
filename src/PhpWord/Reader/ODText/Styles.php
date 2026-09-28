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

namespace PhpOffice\PhpWord\Reader\ODText;

use DOMElement;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Shared\XMLReader;

/**
 * Styles reader: the list styles of styles.xml, and the automatic ones of content.xml.
 */
class Styles extends AbstractPart
{
    /**
     * Word's number format for each `style:num-format` OpenDocument has.
     *
     * @var array<int|string, string>
     */
    private const NUM_FORMATS = ['1' => 'decimal', 'a' => 'lowerLetter', 'A' => 'upperLetter', 'i' => 'lowerRoman', 'I' => 'upperRoman', '' => 'none'];

    /**
     * Read the list styles as numbering styles.
     */
    public function read(PhpWord $phpWord): void
    {
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);

        foreach ($xmlReader->getElements('office:styles/text:list-style|office:automatic-styles/text:list-style') as $listStyle) {
            $levels = [];
            // OpenDocument has ten levels, Word nine
            foreach ($xmlReader->getElements('text:list-level-style-number|text:list-level-style-bullet', $listStyle) as $levelStyle) {
                if ($levelStyle instanceof DOMElement && (int) $levelStyle->getAttribute('text:level') <= 9) {
                    $levels[(int) $levelStyle->getAttribute('text:level') - 1] = $this->readLevel($xmlReader, $levelStyle);
                }
            }
            $phpWord->addNumberingStyle($listStyle->getAttribute('style:name'), ['type' => 'multilevel', 'levels' => $levels]);
        }
    }

    /**
     * Read a level: its number is Word's text such as `%1.%2)`, the numbers of the levels it shows
     * between the prefix and the suffix.
     *
     * @return array<string, int|string>
     */
    private function readLevel(XMLReader $xmlReader, DOMElement $levelStyle): array
    {
        $values = $this->readIndent($xmlReader, $levelStyle);
        $font = $xmlReader->getElement('style:text-properties', $levelStyle);
        if ($font instanceof DOMElement && ($font->getAttribute('style:font-name') ?: $font->getAttribute('fo:font-family')) !== '') {
            $values['font'] = $font->getAttribute('style:font-name') ?: $font->getAttribute('fo:font-family');
        }
        if ($levelStyle->nodeName === 'text:list-level-style-bullet') {
            return ['format' => 'bullet', 'text' => $levelStyle->getAttribute('text:bullet-char')] + $values;
        }
        $level = (int) $levelStyle->getAttribute('text:level');
        $numbers = [];
        for ($shown = max(1, $level - max(1, (int) $levelStyle->getAttribute('text:display-levels')) + 1); $shown <= $level; ++$shown) {
            $numbers[] = '%' . $shown;
        }

        // A level that shows no number keeps only its prefix and suffix. LibreOffice also writes
        // the whole text, `%1%-%2%)`, which keeps separators other than a dot.
        $format = self::NUM_FORMATS[$levelStyle->getAttribute('style:num-format')] ?? 'decimal';
        $text = $levelStyle->getAttribute('style:num-prefix') . ($format === 'none' ? '' : implode('.', $numbers)) . $levelStyle->getAttribute('style:num-suffix');
        if ($format !== 'none' && $levelStyle->getAttribute('loext:num-list-format') !== '') {
            $text = (string) preg_replace('/%(\d+)%/', '%$1', $levelStyle->getAttribute('loext:num-list-format'));
        }

        return [
            'format' => $format,
            'text' => $text,
            'start' => max(1, (int) ($levelStyle->getAttribute('text:start-value') ?: 1)),
        ] + $values;
    }

    /**
     * Read the indent of a level, in either of the two ways OpenDocument has: the position of the
     * label and the text (label-alignment, what Writer writes), or the space before the label and
     * its width.
     *
     * @return array<string, int>
     */
    private function readIndent(XMLReader $xmlReader, DOMElement $levelStyle): array
    {
        $properties = $xmlReader->getElement('style:list-level-properties', $levelStyle);
        if (!$properties instanceof DOMElement) {
            return [];
        }
        $alignment = $xmlReader->getElement('style:list-level-label-alignment', $properties);
        if ($properties->getAttribute('text:list-level-position-and-space-mode') === 'label-alignment' && $alignment instanceof DOMElement) {
            $left = $this->toTwip($alignment->getAttribute('fo:margin-left'));
            $tabPos = $alignment->getAttribute('text:list-tab-stop-position');

            return ['left' => $left, 'hanging' => -$this->toTwip($alignment->getAttribute('fo:text-indent')), 'tabPos' => $tabPos === '' ? $left : $this->toTwip($tabPos)];
        }
        $labelWidth = $this->toTwip($properties->getAttribute('text:min-label-width'));
        $left = $this->toTwip($properties->getAttribute('text:space-before')) + $labelWidth;

        return ['left' => $left, 'hanging' => $labelWidth, 'tabPos' => $left];
    }

    /**
     * A length of OpenDocument in twips.
     */
    private function toTwip(string $length): int
    {
        if (preg_match('/^(-?[\d.]+)(cm|mm|in|pt)$/', $length, $matches) !== 1) {
            return 0;
        }
        $cm = ['cm' => 1, 'mm' => 0.1, 'in' => 2.54, 'pt' => 2.54 / 72][$matches[2]];

        return (int) round(Converter::cmToTwip((float) $matches[1] * $cm));
    }
}
