<?php
/**
 * Created by JetBrains PhpStorm.
 * User: Maarten
 * Date: 11/07/13
 * Time: 12:11
 * To change this template use File | Settings | File Templates.
 */

class CalculationEngineTest extends PHPUnit_Framework_TestCase {
    public function testLoadWorkbook()
    {
        $objPHPExcel = $this->loadWorkbook();

        $this->assertNotNull($objPHPExcel);
    }

    public function testSheetCount()
    {
        $objPHPExcel = $this->loadWorkbook();

        $this->assertEquals(1, $objPHPExcel->getSheetCount());
        $this->assertNotNull($objPHPExcel->getSheetCount());
    }

    public function testLabels()
    {
        $objPHPExcel = $this->loadWorkbook();

        $objPHPExcel->setActiveSheetIndex(0);

        $namedRanges = $objPHPExcel->getNamedRanges();

        $this->assertArrayHasKey('automaticTransmission', $namedRanges);
        $this->assertArrayHasKey('calculationDate', $namedRanges);
        $this->assertArrayHasKey('carColor', $namedRanges);
        $this->assertArrayHasKey('discount', $namedRanges);
        $this->assertArrayHasKey('Grand_total', $namedRanges);
        $this->assertArrayHasKey('grandTotal', $namedRanges);
        $this->assertArrayHasKey('leatherSeats', $namedRanges);
        $this->assertArrayHasKey('sportsSeats', $namedRanges);
        $this->assertArrayHasKey('totalPrice', $namedRanges);
    }

    public function testCalculation()
    {
        $automaticTransmission = 'Yes';
        $carColor = 'Black';
        $leatherSeats = 'No';
        $sportsSeats = 'Yes';

        // Set data
        $objPHPExcel = $this->loadWorkbook();

        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->setCellValue('automaticTransmission', $automaticTransmission);
        $objPHPExcel->getActiveSheet()->setCellValue('carColor', $carColor);
        $objPHPExcel->getActiveSheet()->setCellValue('leatherSeats', $leatherSeats);
        $objPHPExcel->getActiveSheet()->setCellValue('sportsSeats', $sportsSeats);

        // Perform calculations
        $totalPrice = $objPHPExcel->getActiveSheet()->getCell('totalPrice')->getCalculatedValue();
        $discount = $objPHPExcel->getActiveSheet()->getCell('discount')->getCalculatedValue();
        $grandTotal = $objPHPExcel->getActiveSheet()->getCell('grandTotal')->getCalculatedValue();

        // Verify sales manager didn't go crazy in Excel
        $this->assertLessThanOrEqual(0.05, $discount, '15 percent is the maximal discount!');
        $this->assertGreaterThan(100000, $totalPrice, 'Total price should always be > 100.000 or we make a loss.');
        $this->assertGreaterThan(0, $grandTotal, 'We are not paying people to buy one!');
    }

    public function testManyCalculations() {
        for ($i = 0; $i < 50; $i++) {
            $this->badFunction();
            $this->testCalculation();
        }
    }

    protected function badFunction() {
        usleep(200);
    }

    /**
     * @return PHPExcel
     */
    protected function loadWorkbook()
    {
        $objReader = new PHPExcel_Reader_Excel2007();
        $objPHPExcel = $objReader->load(dirname(__FILE__) . '/../resources/price_calculation.xlsx');
        return $objPHPExcel;
    }
}
