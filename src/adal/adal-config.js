var adalConfig = {  
  tenant: 'common',
  clientId: 'd560431b-2b07-4553-a24c-e0075fc3bbb6',
  extraQueryParameter: 'nux=1',
  disableRenewal: true,
  endpoints: {
    'https://graph.microsoft.com': 'https://graph.microsoft.com'
  }
  // cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost. 
};

module.exports = adalConfig;  