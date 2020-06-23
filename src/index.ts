import ssf from 'ssf';
import { Transform, PassThrough, Readable } from 'stream';
import { ReadStream } from 'tty';

import { IWorksheetOptions, IXlsxStreamOptions, IWorksheet } from './types';

const unzip = require('unzip-stream');
const saxStream = require('sax-stream');
const Combine = require('stream-combiner');

function lettersToNumber(letters: string) {
    return letters.split('').reduce((r, a) => r * 26 + parseInt(a, 36) - 9, 0);
}

function numbersToLetter(number: number) {
    let colName = '';
    let dividend = Math.floor(Math.abs(number));
    let rest: number;

    while (dividend > 0) {
        rest = (dividend - 1) % 26;
        colName = String.fromCharCode(65 + rest) + colName;
        dividend = parseInt(`${(dividend - rest) / 26}`);
    }
    return colName;
};


function applyHeaderToObj(obj: any, header: any) {
    if (!header || !header.length) {
        return obj;
    }
    const newObj: { [key: string]: any } = {};
    for (const columnName of Object.keys(obj)) {
        const index = lettersToNumber(columnName) - 1;
        newObj[header[index]] = obj[columnName];
    }
    return newObj;
}

function getTransform(formats: (string | number)[], strings: string[], withHeader?: boolean, ignoreEmpty?: boolean) {
    let lastReceivedRow: number;
    let header: any[] = [];
    return new Transform({
        objectMode: true,
        transform(chunk, encoding, done) {
            let arr: any[] = [];
            let formattedArr = [];
            let obj: any = {};
            let formattedObj: any = {};
            const children = chunk.children ? chunk.children.c.length ? chunk.children.c : [chunk.children.c] : [];
            lastReceivedRow = chunk.attribs.r;
            for (let i = 0; i < children.length; i++) {
                const ch = children[i];
                if (ch.children) {
                    let value: any;
                    if (ch.attribs.t === 'inlineStr') {
                        value = ch.children.is.children.t.value;
                    } else {
                        value = ch.children.v.value;
                        if (ch.attribs.t === 's') {
                            value = strings[value];
                        }
                    }
                    value = isNaN(value) ? value : Number(value);
                    let column = ch.attribs.r.replace(/[0-9]/g, '');
                    const index = lettersToNumber(column) - 1;
                    arr[index] = value;
                    obj[column] = value;
                    const formatId = ch.attribs.s ? Number(ch.attribs.s) : 0;
                    if (formatId) {
                        value = ssf.format(formats[formatId], value);
                        value = isNaN(value) ? value : Number(value);
                    }
                    formattedArr[index] = value;
                    formattedObj[column] = value;
                }
            }
            if (withHeader && !header.length) {
                for (let i = 0; i < arr.length; i++) {
                    const hasDuplicate = arr.filter(x => x === arr[i]).length > 1;
                    header[i] = hasDuplicate ? `[${numbersToLetter(i + 1)}] ${arr[i]}` : arr[i];
                }
                done();
            } else {
                done(undefined, ignoreEmpty && !arr.length ? null : {
                    raw: {
                        obj: applyHeaderToObj(obj, header),
                        arr
                    },
                    formatted: {
                        obj: applyHeaderToObj(formattedObj, header),
                        arr: formattedArr,
                    },
                    header,
                });
            }
        },
        flush(callback) {
            callback();
        }
    })
}

export function getXlsxStream (options: IXlsxStreamOptions): Transform {
    const sheets: string[] = [];
    const numberFormats: any = {};
    const formats: (string | number)[] = [];
    const strings: string[] = [];
    const sheetId = options.sheet + 1
    return Combine(unzip.Parse(), new Transform({
        objectMode: true,
        transform: function(entry, e, cb) {
            const filePath = entry.path;
            switch (filePath) {
                case 'xl/workbook.xml':
                    entry.pipe(saxStream({
                        strict: true,
                        tag: 'sheet'
                    })).on('data', (x: any) => {
                        const attribs = x.attribs;
                        sheets.push(attribs.name);
                    }).on('end', cb);
                    break;
                case 'xl/styles.xml':
                    entry.pipe(saxStream({
                        strict: true,
                        tag: ['cellXfs', 'numFmts']
                    })).on('data', (x: any) => {
                        if (x.tag === 'numFmts' && x.record.children) {
                            const children = x.record.children.numFmt.length ? x.record.children.numFmt : [x.record.children.numFmt];
                            for (let i = 0; i < children.length; i++) {
                                numberFormats[Number(children[i].attribs.numFmtId)] = children[i].attribs.formatCode;
                            }
                        } else if (x.tag === 'cellXfs' && x.record.children) {
                            for (let i = 0; i < x.record.children.xf.length; i++) {
                                const ch = x.record.children.xf[i];
                                formats[i] = Number(ch.attribs.numFmtId);
                            }
                        }
                    }).on('end', () => {
                        for (let i = 0; i < formats.length; i++) {
                            const format = numberFormats[formats[i]];
                            if (format) {
                                formats[i] = format;
                            }
                        }
                        return cb();
                    });
                    break;
                case 'xl/sharedStrings.xml':
                    console.log({ SHAREDSTRINGS: '1' })
                    entry.pipe(saxStream({
                        strict: true,
                        tag: 'si'
                    })).on('data', (x: any) => {
                        if (x.children.t) {
                            strings.push(x.children.t.value);
                        } else if (!x.children.r.length) {
                            strings.push(x.children.r.children.t.value);
                        } else {
                            let str = '';
                            for (let i = 0; i < x.children.r.length; i++) {
                                str += x.children.r[i].children.t.value;
                            }
                            strings.push(str);
                        }
                    }).on('end', cb);
                    break;
                case `xl/worksheets/sheet${sheetId}.xml`:
                    console.log({ SHEET: '1' })
                    const self = this   
                    const pushChunk = function (chunk: any) {
                        if (self.readableLength >= self.readableHighWaterMark) {
                            // Not ready to push now
                            return process.nextTick(pushChunk, [chunk])
                        }
                        return self.push(chunk)
                    }
                    entry
                        .pipe(saxStream({
                            strict: true,
                            tag: 'row'
                        }))
                        .pipe(getTransform(formats, strings, options.withHeader, options.ignoreEmpty))
                        .on('data', (chunk: any) => {
                            return pushChunk(chunk)
                        })
                        .on('end', cb);
                    break;
                default:
                    entry.autodrain();
                    return cb();
                    
            }
        }
    }))
}

export function getWorksheets(options: IWorksheetOptions) {
    const sheets: IWorksheet[] = [];
    return new Promise<IWorksheet[]>((resolve, reject) => {
        options.stream
            .pipe(unzip.Parse())
            .pipe(new Transform({
                objectMode: true,
                transform: function(entry,e,cb) {
                    const filePath = entry.path;
                    if (filePath === 'xl/workbook.xml') {
                        return entry
                            .pipe(saxStream({
                                strict: true,
                                tag: 'sheet'
                            })).on('data', (x: any) => {
                                this.push({
                                    name: x.attribs.name,
                                    hidden: x.attribs.state && x.attribs.state === 'hidden' ? true : false,
                                });
                            }).on('end', cb);
                    }
                    entry.autodrain();
                    return cb();
                }
            }))
            .on('data', (sheet: IWorksheet) => sheets.push(sheet))
            .on('end', () => resolve(sheets))
            .on('error', reject);
    })
}