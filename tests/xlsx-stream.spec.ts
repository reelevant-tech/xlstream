import { getXlsxStream, getWorksheets } from '../src';
import { createReadStream } from 'fs';

it('reads XLSX file correctly', async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/basic.xlsx').pipe(getXlsxStream({
        sheet: 0,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads empty XLSX file correctly', async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/empty.xlsx').pipe(getXlsxStream({
        sheet: 0,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads XLSX file with header', async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/with-header.xlsx').pipe(getXlsxStream({
        sheet: 0,
        withHeader: true,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('reads XLSX file with header values being dupicated', async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/with-header-duplicated.xlsx').pipe(getXlsxStream({
        sheet: 0,
        withHeader: true,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('ignores empty rows', async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/empty-rows.xlsx').pipe(getXlsxStream({
        sheet: 0,
        ignoreEmpty: true,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it('gets worksheets', async (done) => {
    const sheets = await getWorksheets({
        stream: createReadStream('./tests/assets/worksheets.xlsx'),
    });
    expect(sheets).toEqual([
        { name: 'Sheet1', hidden: false, },
        { name: 'Sheet2', hidden: false, },
        { name: 'Sheet3', hidden: false, },
        { name: 'Sheet4', hidden: false, },
    ]);
    done();
});

it('gets worksheets with correct hidden info', async (done) => {
    const sheets = await getWorksheets({
        stream: createReadStream('./tests/assets/hidden-sheet.xlsx'),
    });
    expect(sheets).toEqual([
        { name: 'Sheet1', hidden: false, },
        { name: 'Sheet2', hidden: true, },
        { name: 'Sheet3', hidden: false, },
    ]);
    done();
});

it.only('gets worksheet by index, even if they are reordered', async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/worksheets-reordered.xlsx').pipe(getXlsxStream({
        sheet: 1,
        ignoreEmpty: true,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it(`doesn't fail when empty row has custom height`, async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/empty-row-custom-height.xlsx').pipe(getXlsxStream({
        sheet: 0,
        ignoreEmpty: true,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});

it(`throws expected bad archive error`, async (done) => {
    createReadStream('./tests/assets/bad-archive.xlsx').pipe(getXlsxStream({
        sheet: 0,
    })).on('error', (err) => {
        expect(err).toMatchSnapshot();
        done();
    });
});

it('correctly handles shared string if it has just one value in cell', async (done) => {
    const data: any = [];
    const stream = createReadStream('./tests/assets/shared-string-single-value.xlsx').pipe(getXlsxStream({
        sheet: 0,
        ignoreEmpty: true,
    }));
    stream.on('data', x => data.push(x));
    stream.on('end', () => {
        expect(data).toMatchSnapshot();
        done();
    })
});
