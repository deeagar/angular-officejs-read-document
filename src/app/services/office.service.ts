import { Injectable } from '@angular/core';

declare const Office: any;

@Injectable()
export class OfficeService {

    constructor() { }

    readDocumentFileAsync(): Promise<any> {
        return new Promise((resolve, reject) => {
            const chunkSize = 65536;
            const self = this;

            Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    reject(asyncResult.error);
                } else {
                    self.getAllSlices(asyncResult.value).then(result => {
                        if (result.IsSuccess) {
                            resolve(result.Data);
                        } else {
                            reject(asyncResult.error);
                        }
                    });
                }
            });
        });
    }

    private getAllSlices(file: any): Promise<any> {
        const self = this;
        let isError = false;
        return new Promise(async (resolve, reject) => {
            let documentFileData = [];
            for (let sliceIndex = 0; (sliceIndex < file.sliceCount) && !isError; sliceIndex++) {
                // tslint:disable-next-line:prefer-const
                let sliceReadPromise = new Promise((sliceResolve, sliceReject) => {
                    file.getSliceAsync(sliceIndex, (asyncResult) => {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            documentFileData = documentFileData.concat(asyncResult.value.data);
                            sliceResolve({
                                IsSuccess: true,
                                Data: documentFileData
                            });
                        } else {
                            file.closeAsync();
                            sliceReject({
                                IsSuccess: false,
                                ErrorMessage: `Error in reading the slice: ${sliceIndex} of the document`
                            });
                        }
                    });
                });
                await sliceReadPromise.catch((error) => {
                    isError = true;
                });
            }

            if (isError || !documentFileData.length) {
                reject('Error while reading document. Please try it again.');
                return;
            }

            const encodedFileString = self.encodeBase64(documentFileData);
            file.closeAsync();

            resolve({
                IsSuccess: true,
                Data: encodedFileString
            });
        });
    }

    encodeBase64(documentFileData) {
        let data = '';

        // tslint:disable-next-line:prefer-for-of
        for (let i = 0; i < documentFileData.length; i++) {
            data += String.fromCharCode(documentFileData[i]);
        }
        if (data) {
            return window.btoa(data);
        }
        return null;
    }
}
