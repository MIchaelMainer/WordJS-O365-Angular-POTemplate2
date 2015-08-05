/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

(function () {
	angular
		.module('poTemplateApp')
		.controller('NavbarController', NavbarController);
 
	/**
	 * The NavbarController code.
	 */
	function NavbarController($log, adalAuthenticationService) {
		var vm = this;
		
		// Properties
		vm.isCollapsed;
		
		// Methods
		vm.connect = connect;
		vm.disconnect = disconnect;
		 
		/////////////////////////////////////////
		// End of exposed properties and methods.
		
		// Activate controller when it loads.
		activate();
		
		/**
		 * This function does any initialization work the 
		 * controller needs.
		 */
		function activate() {
			vm.isCollapsed = true;
		};
		
		/**
		 * Expose the login method to the view.
		 */
		function connect() {
			$log.debug('Connecting to Office 365...');
			adalAuthenticationService.login();
		};
		
		/**
		 * Expose the logOut method to the view.
		 */
		function disconnect() {
			$log.debug('Disconnecting from Office 365...');
			adalAuthenticationService.logOut();
		};
	};
})();

