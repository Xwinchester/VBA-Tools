function BuildNavbar() {
    // Array of navigation links
    const navLinks = [
        { text: 'Class Generator', url: 'class_generator.html' },
        { text: 'Snippit Manager', url: 'snippits.html' }
        // Add more links as needed
    ];

    // Function to build the navigation links
    function buildNav() {
        const navContainer = document.getElementById('nav-links');
        
        // Clear any existing links
        navContainer.innerHTML = '';

        // Iterate through the links array and create list items
        navLinks.forEach(link => {
            const listItem = document.createElement('li');
            const anchor = document.createElement('a');
            anchor.href = link.url;
            anchor.textContent = link.text;
            anchor.classList.add('nav-link'); // Add a class for styling
            
            // Append the anchor to the list item
            listItem.appendChild(anchor);
            // Append the list item to the navigation container
            navContainer.appendChild(listItem);
        });
    }

    // Call the buildNav function to execute it
    buildNav();
}

function BuildFooter() {
    // Function to create and insert the footer
    function createFooter() {
        const footer = document.createElement('footer');
        footer.classList.add('footer');

        const developerText = document.createElement('p');
        developerText.classList.add('footer-text'); 
        developerText.innerHTML = `Developed by <strong>Drew Winchester</strong>`;
        footer.appendChild(developerText);

        // Append the footer to the specified div or body
        document.getElementById('footer').appendChild(footer);
    }

    // Call the function to create the footer
    createFooter();
}

// Call the functions to build the navigation and footer on page load
window.onload = function() {
    BuildNavbar();
    BuildFooter();
};
