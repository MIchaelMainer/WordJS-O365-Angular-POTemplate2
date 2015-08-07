/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
	angular
		.module('poTemplateApp')
		.factory('office365Factory', office365Factory);

	function office365Factory($log, $http) {
		var office365 = {}; 
 
		// Methods
        office365.getContactList = getContactList;
        office365.getSelectedContact = getSelectedContact;
        
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
//        $http.defaults.useXDomain = true;
//        delete $http.defaults.headers.common['X-Requested-With'];
        
        baseUrl = 'https://outlook.office365.com/api/v1.0/me/contacts';
        
        function getContactList() {
            var request = {
                method: 'GET',
                url: baseUrl + '?$select=GivenName,Surname&$orderby=surname'
            };
            
            return $http(request);
        };
        
        function getSelectedContact(id) {
            var request = {
                method: 'GET',
                url: baseUrl + '/' + id + '?$select=GivenName,Surname,' +
                     'BusinessAddress,BusinessPhones,CompanyName,EmailAddresses'
            };
            
            return $http(request);
        };        
     

		return office365;
	};
})();

