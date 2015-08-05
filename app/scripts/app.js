/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
  angular
    .module('poTemplateApp', [
      'ngRoute',
      'AdalAngular',
      'ui.bootstrap',
      'angular-loading-bar'
    ])
    .config(config);

  function config($routeProvider, $httpProvider, adalAuthenticationServiceProvider, cfpLoadingBarProvider) {
    // Configure the routes. 
    $routeProvider
      .when('/', {
        templateUrl: 'views/main.html',
        controller: 'MainController',
        controllerAs: 'main',
        requireADLogin: true
      })
      .otherwise({
        redirectTo: '/'
      });
      
    // Configure ADAL JS. 
    adalAuthenticationServiceProvider.init(
      {
        tenant: tenant,
        clientId: clientId,
        //redirectUri: 'http://127.0.0.1:8080',
        //cacheLocation: 'localStorage',
        endpoints: {
            'https://outlook.office365.com': 'https://outlook.office365.com'
        }
        
      },
      $httpProvider
      );
    
    // Loading bar configuration options.
    cfpLoadingBarProvider.includeSpinner = false;
  };
})(); 


