import * as React from 'react';
import { render } from 'react-dom';
import { App } from './components/app';
import { Progress } from './components/progress';
import './assets/styles/global.scss';
import Adal = require('./adal/adal-request.js');

(() => {
    const title = 'My Office Add-in';
    const container = document.querySelector('#container');
    let component = this;
    Adal.processAdalCallback();
    /* Render application after Office initializes */
    Office.initialize = () => {
        if (window === window.parent) {
            component.serverRequest = Adal.adalRequest({
                url: 'https://graph.microsoft.com/v1.0/me/memberOf?$top=500',
                headers: {
                    'Accept': 'application/json;odata.metadata=full'
                }
            }).then((data) => {
                console.log(data);
            });
        };


        render(
            <App title={title} />,
            container
        );
    };

    if (window === window.parent) {
        debugger;
        component.serverRequest = Adal.adalRequest({
            url: 'https://graph.microsoft.com/v1.0/me/memberOf?$top=500',
            headers: {
                'Accept': 'application/json;odata.metadata=full'
            }
        }).then((data) => {
            debugger;
            console.log(data);
        });
        /* Initial render showing a progress bar */
        render(<Progress title={title} logo='assets/logo-filled.png' message='Please sideload your addin to see app body.' />, container);
    }
})();
