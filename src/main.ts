import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

declare const Office: any;
if (environment.production) {
    enableProdMode();
}

function launch() {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .then((success: any) => {
            console.log('Main: AppModule bootstrap success', success);
        })
        .catch((error: any) => {
            console.log('Main: AppModule bootstrap error', error);
        });
}

Office.onReady(() => {
    console.log('Main - Initializing office.js');
    launch();
    console.log('Main - Office.js is initialized');
});
