/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
	angular
		.module('poTemplateApp')
		.controller('MainController', MainController);

	MainController.$inject = ['$log', '$q', 'adalAuthenticationService', 'office365Factory'];
    
	/**
	 * The MainController code.
	 */
	function MainController($log, $q, adalAuthenticationService, office365) {
		
        // Setup the viewmodel
        var vm = this;
		
        // properties
        vm.contacts = [];
        
        // methods
        vm.getContactList = getContactList;
        vm.getSelectedContact = getSelectedContact;
            
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		// Activate controller when it loads. Should turn this into an IIFE
        activate();
        
        function activate() {
			// Once the user is logged in, fetch the data.
			if (adalAuthenticationService.userInfo.isAuthenticated) {
				getContactList();
			}
		};
        
        function getContactList() {
			return $q(function (resolve, reject) {
				office365.getContactList()
					.then(function (response) {
                        
                        // Bind the contacts in the response to the view model.
                        setupContactViewModel(response);
                    
                        resolve();
					}, function (err) {
						reject(err);
					});
			});
		};
        
        // Update the view model to display lastname, firstname.
        // Add the contacts to the viewmodel.
        function setupContactViewModel(response) {
            
            var contacts = response.data.value;
            
            if (contacts.length > 0)
            {
                for (i = 0; i < contacts.length; i++) {
                    contacts[i].displayName = contacts[i].Surname + 
                                              ', ' + 
                                              contacts[i].GivenName;
                }     
            }
            
            vm.contacts = contacts;
        };
        
        function getSelectedContact(id) {
            
            // In case of null selection.
            if (id == null)
                return; 
            
            $log.log('Selected contact: ' + id);
            
			return $q(function (resolve, reject) {
				office365.getSelectedContact(id)
					.then(function (response) {
                        
                        $log.log('Contact info: ' + JSON.stringify(response.data) );
                        
                        // Send the selected contact information to 
                        // be inserted into Word.
                        insertContact(response.data);
                    
                    
						resolve();
					}, function (err) {
						reject(err);
					});
			});
		};
        
        function insertContact(contact){

            // Prepare info for the view (Word)
            var contactName = contact.Surname + ', ' + contact.GivenName;
            var street = contact.BusinessAddress.Street;
            var cityStateZIP = contact.BusinessAddress.City + ", " +
                               contact.BusinessAddress.State + " " +
                               contact.BusinessAddress.PostalCode;
            var businessPhone1 = contact.BusinessPhones[0];
            
            // Get the client request context.
            var ctx = new Word.RequestContext();

            var contentControls = ctx.document.body.contentControls;

            ctx.load(contentControls, {select: 'text,tag,id',
                                       expand: 'items'});

            // Run the set of actions in the queue. 
            ctx.executeAsync()
                .then(function () {
                    for (var i = 0; i < contentControls.items.length; i++) {

                        // Set the text value of the content controls.
                        switch(contentControls.items[i].tag) {
                            case 'purchaserName':
                                contentControls.items[i].removeWhenEdited = false;
                                contentControls.items[i].insertText(contactName, Word.InsertLocation.replace);
                                break;
                            case 'companyName':
                                contentControls.items[i].removeWhenEdited = false;
                                contentControls.items[i].insertText(contact.CompanyName, Word.InsertLocation.replace);
                                break;
                            case 'streetAddress':
                                contentControls.items[i].removeWhenEdited = false;
                                contentControls.items[i].insertText(street, Word.InsertLocation.replace);
                                break;
                            case 'citySTzip':
                                contentControls.items[i].removeWhenEdited = false;
                                contentControls.items[i].insertText(cityStateZIP, Word.InsertLocation.replace);
                                break;
                            case 'phone':
                                contentControls.items[i].removeWhenEdited = false;
                                contentControls.items[i].insertText(businessPhone1, Word.InsertLocation.replace);
                                break;
                            default: break;
                        }
                    }
                
                    ctx.executeAsync()
                        .then(function(){
                            console.log("Success");
                        });
            
                })
                .catch(function(error){
                    console.log("ERROR: " + JSON.stringify(error));
                });
        }
	};
})();

