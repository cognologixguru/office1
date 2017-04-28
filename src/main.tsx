import * as React from 'react';
import { render } from 'react-dom';
import { App } from './components/app';
import { Progress } from './components/progress';
import './assets/styles/global.scss';
import Adal = require('./adal/adal-request.js');

(() => {
    const title = 'My Office Add-in';
    const container = document.querySelector('#container');

    /* Render application after Office initializes */
    Office.initialize = () => {
        Adal.processAdalCallback();

        render(
            <App title={title} />,
            container
        );
    }
});

Adal.processAdalCallback();
if (window === window.parent) {
    /* Initial render showing a progress bar */
    render(<Progress title={title} logo='assets/logo-filled.png' message='Please sideload your addin to see app body.' />, container);
}
})();

