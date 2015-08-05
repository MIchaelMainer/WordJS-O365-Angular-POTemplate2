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
		
        // prop
        vm.selectedContact;
        vm.contacts = [];
        
        // methods
        vm.getContactList = getContactList;
        vm.getSelectedContact = getSelectedContact;
        
        
        
        
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		// Activate controller when it loads.
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
						// Bind data to the view model.
						vm.contacts = response.data.value;	
                        $log.log('We got back ' + vm.contacts.length + ' contact(s)' );
                    
						resolve();
					}, function (err) {
						reject(err);
					});
			});
		};
        
        function getSelectedContact(id) {
            
            // In case of null selection.
            if (id == null)
                return; 
            
            $log.log('Selected contact: ' + id);
            
			return $q(function (resolve, reject) {
				office365.getSelectedContact(id)
					.then(function (response) {
						// Bind data to the view model.
						vm.selectedContact = response.data.value;	

                        InsertContact(vm.selectedContact);
                    
                    
						resolve();
					}, function (err) {
						reject(err);
					});
			});
		};
        
        
        function InsertContact(contact){

            var ctx = new Word.RequestContext();

            // Queue: get the user's current selection and create a range object named range.
            // Queue: insert 'Hello World!' at the end of the selection.
            var range = ctx.document.getSelection();
            
            range.insertText('Inserted', Word.InsertLocation.end);
//            range.insertText(JSON.stringify(contact), Word.InsertLocation.end);

            // Run the set of actions in the queue. In this case, we are inserting text
            // at the end of range. 
            ctx.executeAsync()
                .then(function () {
                    console.log("Done");
                })
                .catch(function(error){
                    console.log("ERROR: " + JSON.stringify(error));
                });
        }
	};
})();

