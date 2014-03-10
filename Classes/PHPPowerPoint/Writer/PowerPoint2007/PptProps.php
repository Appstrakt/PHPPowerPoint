<?php
/**
 * PHPPowerPoint
 *
 * Copyright (c) 2009 - 2010 PHPPowerPoint
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    0.1.0, 2009-04-27
 */


/** PHPPowerPoint */
require_once __DIR__.'/../../../PHPPowerPoint.php';

/** PHPPowerPoint_Writer_PowerPoint2007 */
require_once __DIR__.'/../PowerPoint2007.php';

/** PHPPowerPoint_Writer_PowerPoint2007_WriterPart */
require_once __DIR__.'/WriterPart.php';

/** PHPPowerPoint_Shared_XMLWriter */
require_once __DIR__.'/../../Shared/XMLWriter.php';


/**
 * PHPPowerPoint_Writer_PowerPoint2007_PptProps
 *
 * @category   PHPPowerPoint
 * @package    PHPPowerPoint_Writer_PowerPoint2007
 * @copyright  Copyright (c) 2009 - 2010 PHPPowerPoint (http://www.codeplex.com/PHPPowerPoint)
 */
class PHPPowerPoint_Writer_PowerPoint2007_PptProps extends PHPPowerPoint_Writer_PowerPoint2007_WriterPart
{
/**
     * Write ppt/presProps.xml to XML format
     *
     * @param   PHPPowerPoint   $pPHPPowerPoint
     * @return  string      XML Output
     * @throws  Exception
     */
    public function writePresProps(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0','UTF-8','yes');

        // p:presentationPr
        $objWriter->startElement('p:presentationPr');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

            // p:extLst
            $objWriter->startElement('p:extLst');

                // p:ext
                $objWriter->startElement('p:ext');
                $objWriter->writeAttribute('uri', '{E76CE94A-603C-4142-B9EB-6D1370010A27}');

                    // p14:discardImageEditData
                    $objWriter->startElement('p14:discardImageEditData');
                    $objWriter->writeAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
                    $objWriter->writeAttribute('val',   '0');
                    $objWriter->endElement();

                $objWriter->endElement();

                // p:ext
                $objWriter->startElement('p:ext');
                $objWriter->writeAttribute('uri', '{D31A062A-798A-4329-ABDD-BBA856620510}');

                    // p14:defaultImageDpi
                    $objWriter->startElement('p14:defaultImageDpi');
                    $objWriter->writeAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
                    $objWriter->writeAttribute('val',   '220');
                    $objWriter->endElement();

                $objWriter->endElement();

            $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write ppt/tableStyles.xml to XML format
     *
     * @param   PHPPowerPoint   $pPHPPowerPoint
     * @return  string      XML Output
     * @throws  Exception
     */
    public function writeTableStyles(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0','UTF-8','yes');

        // a:tblStyleLst
        $objWriter->startElement('a:tblStyleLst');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('def', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }

    /**
     * Write ppt/viewProps.xml to XML format
     *
     * @param   PHPPowerPoint   $pPHPPowerPoint
     * @return  string      XML Output
     * @throws  Exception
     */
    public function writeViewProps(PHPPowerPoint $pPHPPowerPoint = null)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new PHPPowerPoint_Shared_XMLWriter(PHPPowerPoint_Shared_XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0','UTF-8','yes');

        // p:viewPr
        $objWriter->startElement('p:viewPr');
        $objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
        $objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $objWriter->writeAttribute('xmlns:p', 'http://schemas.openxmlformats.org/presentationml/2006/main');

            // p:normalViewPr
            $objWriter->startElement('p:normalViewPr');

                // p:restoredLeft
                $objWriter->startElement('p:restoredLeft');
                $objWriter->writeAttribute('sz', '15620');
                $objWriter->endElement();
                // p:restoredTop
                $objWriter->startElement('p:restoredTop');
                $objWriter->writeAttribute('sz', '94660');
                $objWriter->endElement();

            $objWriter->endElement();

            // p:slideViewPr
            $objWriter->startElement('p:slideViewPr');

                // p:cSldViewPr
                $objWriter->startElement('p:cSldViewPr');

                    // p:cViewPr
                    $objWriter->startElement('p:cViewPr');
                    $objWriter->writeAttribute('varScale', '1');

                        // p:scale
                        $objWriter->startElement('p:scale');

                            // a:sx
                            $objWriter->startElement('a:sx');
                            $objWriter->writeAttribute('n', '70');
                            $objWriter->writeAttribute('d', '100');
                            $objWriter->endElement();

                            // a:sy
                            $objWriter->startElement('a:sy');
                            $objWriter->writeAttribute('n', '70');
                            $objWriter->writeAttribute('d', '100');
                            $objWriter->endElement();

                        $objWriter->endElement();

                        // p:origin
                        $objWriter->startElement('p:origin');
                        $objWriter->writeAttribute('x', '-516');
                        $objWriter->writeAttribute('y', '-90');
                        $objWriter->endElement();

                    $objWriter->endElement();

                    // p:guideLst
                    $objWriter->startElement('p:guideLst');

                        // p:guide
                        $objWriter->startElement('p:guide');
                        $objWriter->writeAttribute('orient', 'horz');
                        $objWriter->writeAttribute('pos', '2160');
                        $objWriter->endElement();

                        // p:guide
                        $objWriter->startElement('p:guide');
                        $objWriter->writeAttribute('pos', '2880');
                        $objWriter->endElement();

                    $objWriter->endElement();

                $objWriter->endElement();

            $objWriter->endElement();

            // p:notesTextViewPr
            $objWriter->startElement('p:notesTextViewPr');

                // p:cViewPr
                $objWriter->startElement('p:cViewPr');

                    // p:scale
                    $objWriter->startElement('p:scale');

                        // a:sx
                        $objWriter->startElement('a:sx');
                        $objWriter->writeAttribute('n', '1');
                        $objWriter->writeAttribute('d', '1');
                        $objWriter->endElement();

                        // a:sy
                        $objWriter->startElement('a:sy');
                        $objWriter->writeAttribute('n', '1');
                        $objWriter->writeAttribute('d', '1');
                        $objWriter->endElement();

                    $objWriter->endElement();

                    // p:origin
                    $objWriter->startElement('p:origin');
                    $objWriter->writeAttribute('x', '0');
                    $objWriter->writeAttribute('y', '0');
                    $objWriter->endElement();

                $objWriter->endElement();

            $objWriter->endElement();

            // p:gridSpacing
            $objWriter->startElement('p:gridSpacing');
            $objWriter->writeAttribute('cx', '76200');
            $objWriter->writeAttribute('cy', '76200');
            $objWriter->endElement();

        $objWriter->endElement();

        // Return
        return $objWriter->getData();
    }
}