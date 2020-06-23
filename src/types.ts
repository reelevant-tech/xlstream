import { Readable } from 'stream'

export interface IXlsxStreamOptions {
    sheet: number;
    withHeader?: boolean;
    ignoreEmpty?: boolean;
}

export interface IWorksheetOptions {
    stream: Readable;
}

export interface IWorksheet {
    name: string;
    hidden: boolean;
}