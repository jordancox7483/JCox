// ==UserScript==
// @name         Remove AOL Mail Section
// @namespace    https://mail.aol.com/
// @version      1.0
// @description  Remove specific section from AOL Mail
// @author       Your Name
// @match        https://mail.aol.com/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Function to remove the specified section
    function removeSection() {
        const sectionSelector = '#mail-app-component-container > div.D_F.o_v.p_R.I_T > div.D_F.ek_BB';
        const sectionElement = document.querySelector(sectionSelector);
        if (sectionElement) {
            sectionElement.remove();
        }
    }

    // Run the function when the page loads
    window.addEventListener('load', removeSection);
})();
