const puppeteer = require('puppeteer');

// Convert using puppeteer
// result is the response from a request to /querydata
const extract = async (result) => {
    const browser = await puppeteer.launch({headless: true});
    const page = await browser.newPage();
    await page.goto('https://app.powerbi.com/view');

    console.log('Waiting for PowerBI initialization...');
    await page.waitForFunction(() => 'powerbi' in window);

    console.log('Extracting data...');
    const timeStart = new Date();
    const evaluateResult = await page.evaluate((res) => {
        const PBIDSR = powerbi.data.dsr;
        const dsrParserContext = PBIDSR.createDsrParserContext(res.descriptor);
        const dsrParser = PBIDSR.reader.V2.createDsrParser(res.dsr, dsrParserContext);
        const converterContext = new PBIDSR.DataViewConverterContext(8, res.descriptor, dsrParserContext);
        const metaColumns = {columns: converterContext.createDataViewMetadataColumns()};
        return PBIDSR.converters.createTable(converterContext, dsrParser, metaColumns);
    }, result);

    const timeEnd = new Date();
    browser.close();

    console.log(`Done extracting (${evaluateResult.rows.length} rows - ${timeEnd-timeStart}ms)`);

    return evaluateResult;
}

module.exports = extract;
