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

use DateTime;
use DOMElement;
use DOMNodeList;
use PhpOffice\Math\Reader\MathML;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\Element\TrackChange;
use PhpOffice\PhpWord\Exception\InvalidImageException;
use PhpOffice\PhpWord\Exception\UnsupportedImageTypeException;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Shared\XMLReader;
use PhpOffice\PhpWord\Style\Image as ImageStyle;

/**
 * Content reader.
 *
 * @since 0.10.0
 */
class Content extends AbstractPart
{
    /** @var ?Section */
    private $section;

    /**
     * Read content.xml.
     */
    public function read(PhpWord $phpWord): void
    {
        $xmlReader = new XMLReader();
        $xmlReader->getDomFromZip($this->docFile, $this->xmlFile);

        $nodes = $xmlReader->getElements('office:body/office:text/*');
        $this->section = null;
        $this->processNodes($nodes, $xmlReader, $phpWord);
        $this->section = null;
    }

    /** @param DOMNodeList<DOMElement> $nodes */
    public function processNodes(DOMNodeList $nodes, XMLReader $xmlReader, PhpWord $phpWord): void
    {
        if ($nodes->length > 0) {
            foreach ($nodes as $node) {
                // $styleName = $xmlReader->getAttribute('text:style-name', $node);
                switch ($node->nodeName) {
                    case 'text:h': // Heading
                        $depth = $xmlReader->getAttribute('text:outline-level', $node);
                        $this->getSection($phpWord)->addTitle($node->nodeValue, $depth);

                        break;
                    case 'text:p': // Paragraph
                        $styleName = $xmlReader->getAttribute('text:style-name', $node);
                        if (substr((string) $styleName, 0, 2) === 'SB') {
                            break;
                        }
                        $element = $xmlReader->getElement('draw:frame/draw:object', $node);
                        if ($element) {
                            $mathFile = str_replace('./', '', $element->getAttribute('xlink:href')) . '/content.xml';

                            $xmlReaderObject = new XMLReader();
                            $mathElement = $xmlReaderObject->getDomFromZip($this->docFile, $mathFile);
                            if ($mathElement) {
                                $mathXML = $mathElement->saveXML($mathElement);

                                if (is_string($mathXML)) {
                                    $reader = new MathML();
                                    $math = $reader->read($mathXML);

                                    $this->getSection($phpWord)->addFormula($math);
                                }
                            }
                        } else {
                            $children = $node->childNodes;
                            $spans = false;
                            /** @var DOMElement $child */
                            foreach ($children as $child) {
                                switch ($child->nodeName) {
                                    case 'text:change-start':
                                        $changeId = $child->getAttribute('text:change-id');
                                        if (isset($trackedChanges[$changeId])) {
                                            $changed = $trackedChanges[$changeId];
                                        }

                                        break;
                                    case 'text:change-end':
                                        unset($changed);

                                        break;
                                    case 'text:change':
                                        $changeId = $child->getAttribute('text:change-id');
                                        if (isset($trackedChanges[$changeId])) {
                                            $changed = $trackedChanges[$changeId];
                                        }

                                        break;
                                    case 'text:span':
                                        $spans = true;

                                        break;
                                }
                            }

                            // One query per prefix: a prefix the document does not declare fails the whole query
                            if ($spans || $xmlReader->elementExists('.//draw:image', $node) || $xmlReader->elementExists('.//text:a', $node)) {
                                $element = $this->getSection($phpWord)->addTextRun();
                                $this->readTextRun($xmlReader, $node, $element);
                            } else {
                                $element = $this->getSection($phpWord)->addText($node->nodeValue);
                            }
                            if (isset($changed) && is_array($changed)) {
                                $element->setTrackChange($changed['changed']);
                                if (isset($changed['textNodes'])) {
                                    foreach ($changed['textNodes'] as $changedNode) {
                                        $element = $this->getSection($phpWord)->addText($changedNode->nodeValue);
                                        $element->setTrackChange($changed['changed']);
                                    }
                                }
                            }
                        }

                        break;
                    case 'text:list': // List
                        $listItems = $xmlReader->getElements('text:list-item/text:p', $node);
                        foreach ($listItems as $listItem) {
                            // $listStyleName = $xmlReader->getAttribute('text:style-name', $listItem);
                            $this->getSection($phpWord)->addListItem($listItem->nodeValue, 0);
                        }

                        break;
                    case 'text:tracked-changes':
                        $changedRegions = $xmlReader->getElements('text:changed-region', $node);
                        foreach ($changedRegions as $changedRegion) {
                            $type = ($changedRegion->firstChild->nodeName == 'text:insertion') ? TrackChange::INSERTED : TrackChange::DELETED;
                            $creatorNode = $xmlReader->getElements('office:change-info/dc:creator', $changedRegion->firstChild);
                            $author = $creatorNode[0]->nodeValue;
                            $dateNode = $xmlReader->getElements('office:change-info/dc:date', $changedRegion->firstChild);
                            $date = $dateNode[0]->nodeValue;
                            $date = preg_replace('/\.\d+$/', '', $date);
                            $date = DateTime::createFromFormat('Y-m-d\TH:i:s', $date);
                            $changed = new TrackChange($type, $author, $date);
                            $textNodes = $xmlReader->getElements('text:deletion/text:p', $changedRegion);
                            $trackedChanges[$changedRegion->getAttribute('text:id')] = ['changed' => $changed, 'textNodes' => $textNodes];
                        }

                        break;
                    case 'text:section': // Section
                        // $sectionStyleName = $xmlReader->getAttribute('text:style-name', $listItem);
                        $this->section = $phpWord->addSection();
                        /** @var DOMNodeList<DOMElement> $children */
                        $children = $node->childNodes;
                        $this->processNodes($children, $xmlReader, $phpWord);

                        break;
                }
            }
        }
    }

    /**
     * Read the text and the images of a paragraph into a run, in their order.
     */
    private function readTextRun(XMLReader $xmlReader, DOMElement $node, TextRun $run): void
    {
        /** @var DOMElement $child */
        foreach ($node->childNodes as $child) {
            switch ($child->nodeName) {
                case '#text':
                    $run->addText($child->nodeValue);

                    break;
                case 'text:tab':
                    $run->addText("\t");

                    break;
                case 'text:s':
                    $spaces = (int) $child->getAttribute('text:c') ?: 1;
                    $run->addText(str_repeat(' ', $spaces));

                    break;
                case 'text:line-break':
                    $run->addTextBreak();

                    break;
                case 'text:note':
                    // The text of a note is not the text of the paragraph

                    break;
                case 'draw:frame':
                    if (!$this->readFrameImage($xmlReader, $child, $run)) {
                        // A caption holds its image and its text in a text box
                        foreach ($xmlReader->getElements('draw:text-box', $child) as $textBox) {
                            $this->readTextRun($xmlReader, $textBox, $run);
                        }
                    }

                    break;
                case 'draw:a':
                    $this->readTextRun($xmlReader, $child, $run);

                    break;
                case 'text:a':
                    $href = $child->getAttribute('xlink:href');
                    // LibreOffice reads a link without a target as its text.
                    // ponytail: a Link holds only text, so a link around an image or another link is read as its
                    // content and loses its target; split the link around them if that matters
                    if ($href === '' || $xmlReader->elementExists('.//text:a', $child) || $xmlReader->elementExists('.//draw:frame', $child)) {
                        $this->readTextRun($xmlReader, $child, $run);

                        break;
                    }
                    $text = new TextRun();
                    $this->readTextRun($xmlReader, $child, $text);
                    // A target in the document, such as a bookmark, starts with #
                    $internal = $href[0] === '#';
                    $run->addLink($internal ? substr($href, 1) : $href, $text->getText(), null, null, $internal);

                    break;
                default:
                    // Spans, fields and the paragraphs of a text box: their text
                    if (strpos($child->nodeName, 'text:') === 0) {
                        $this->readTextRun($xmlReader, $child, $run);
                    }
            }
        }
    }

    /**
     * Read the image of a frame; false if the frame holds none.
     */
    private function readFrameImage(XMLReader $xmlReader, DOMElement $frame, TextRun $run): bool
    {
        $images = $xmlReader->getElements('draw:image', $frame);
        // The image of an object, such as a formula, is only its replacement
        if ($images->length === 0 || $xmlReader->elementExists('draw:object', $frame)) {
            return false;
        }
        if (!$this->hasImageLoading()) {
            return true;
        }

        $style = ['unit' => ImageStyle::UNIT_PX];
        foreach (['width', 'height'] as $key) {
            $length = $frame->getAttribute('svg:' . $key);
            if (preg_match('/^(\d+\.?\d*|\.\d+)(cm|mm|in|pt|pc|px)$/', $length) === 1) {
                // ODF allows a length without its leading zero, which the converter does not read
                $style[$key] = Converter::pointToPixel((float) Converter::cssToPoint((string) preg_replace('/^\./', '0.', $length)));
            }
        }
        // As LibreOffice gives it to a PDF: the title and the description, either alone if the other is empty
        $altText = implode(' - ', array_filter([
            (string) $xmlReader->getValue('svg:title', $frame),
            (string) $xmlReader->getValue('svg:desc', $frame),
        ], static function (string $text): bool {
            return $text !== '';
        }));
        $name = $frame->getAttribute('draw:name');

        // The first image this reader can load; the others are its fallbacks
        foreach ($images as $image) {
            $path = (string) preg_replace('#^\./#', '', rawurldecode($image->getAttribute('xlink:href')));
            // A picture of the package only, never a file or a URL the document points at
            if ($path === '' || $path[0] === '/' || strpos($path, ':') !== false || in_array('..', explode('/', $path), true)) {
                continue;
            }

            try {
                $run->addImage("zip://{$this->docFile}#{$path}", $style, false, $name === '' ? null : $name, $altText === '' ? null : $altText);

                return true;
            } catch (InvalidImageException|UnsupportedImageTypeException $exception) {
                continue;
            }
        }

        return true;
    }

    private function getSection(PhpWord $phpWord): Section
    {
        $section = $this->section;
        if ($section === null) {
            $section = $this->section = $phpWord->addSection();
        }

        return $section;
    }
}
