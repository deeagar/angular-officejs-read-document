import { Component } from '@angular/core';
import { OfficeService } from './services/office.service';

declare const Word: any;

@Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
})
export class AppComponent {
    constructor(private officeService: OfficeService) {

    }

    title = 'read-word-document';
    readDocument(): void {
        this.officeService.readDocumentFileAsync()
            .then((result) => {
                Word.run((context) => {
                    const createdDoc = context.application.createDocument(result);
                    context.load(createdDoc);
                    return context.sync()
                        .then(() => {
                            createdDoc.open();
                            context.sync();
                        }).catch((error) => {
                            console.log(JSON.parse(error));
                        });
                });
            })
            .catch((error) => {
                console.log(JSON.parse(error));
            });
    }
}
