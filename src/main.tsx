import * as React from 'react';
import { render } from 'react-dom';
import { App } from './components/app';
import { Progress } from './components/progress';
import './assets/styles/global.scss';
import { Authenticator, DefaultEndpoints, Utilities, IToken } from '@microsoft/office-js-helpers';


let authenticator = new Authenticator();

function getToken(): IToken {
    return authenticator.tokens.get(DefaultEndpoints.AzureAD);
}

// register Microsoft (Azure AD 2.0 Converged auth) endpoint using

authenticator.endpoints.registerAzureADAuth('d560431b-2b07-4553-a24c-e0075fc3bbb6', 'sanitariumdev.onmicrosoft.com');
debugger;
// for the default Microsoft endpoint
authenticator
    .authenticate(DefaultEndpoints.AzureAD)
    .then((token) => { /* Microsoft Token */
        debugger;
        console.log(token);
    })
    .catch(Utilities.log);



(() => {
    const title = 'My Office Add-in';
    const container = document.querySelector('#container');

    /* Render application after Office initializes */
    Office.initialize = () => {
        if (Authenticator.isAuthDialog()) {
            return;
        };
        render(
            <App title={title} getToken={getToken} />,
            container
        );

    };

    /* Initial render showing a progress bar */
    render(<Progress title={title} logo='assets/logo-filled.png' message='Please sideload your addin to see app body.' />, container);

})();
