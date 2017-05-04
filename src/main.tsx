import * as React from 'react';
import { render } from 'react-dom';
import { App } from './components/app';
import { Progress } from './components/progress';
import './assets/styles/global.scss';
import { Authenticator } from '@microsoft/office-js-helpers';
import { PlannerModel } from './core';


// let authenticator = new Authenticator();

// register Microsoft (Azure AD 2.0 Converged auth) endpoint using


debugger;
let model = new PlannerModel();

// for the default Microsoft endpoint
// authenticator
//     .authenticate(DefaultEndpoints.AzureAD)
//     .then((token) => { /* Microsoft Token */
//         debugger;
//         console.log(token);
//     })
//     .catch(Utilities.log);


(() => {
    const title = 'My Office Add-in';
    const container = document.querySelector('#container');

    /* Render application after Office initializes */
    Office.initialize = () => {
        if (!Authenticator.isAuthDialog()) {
            render(
                <App title={title} />,
                container
            );
        };
    };

    /* Initial render showing a progress bar */
    render(<Progress title={title} logo='assets/logo-filled.png' message='Please sideload your addin to see app body.' />, container);

})();
