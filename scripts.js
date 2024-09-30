// Function to load content dynamically from other HTML files
function loadContent(page, element) {
  const xhttp = new XMLHttpRequest();
  xhttp.onload = function () {
    document.getElementById("page-content").innerHTML = this.responseText;
  }
  xhttp.open("GET", page, true);
  xhttp.send();

  // Remove 'active' class from all menu links
  const menuLinks = document.getElementsByClassName("menu-link");
  for (let i = 0; i < menuLinks.length; i++) {
    menuLinks[i].classList.remove("active");
  }

  // Add 'active' class to the clicked link
  element.classList.add("active");
}

// Function to toggle the submenu visibility
function toggleMenu(header) {
  const submenu = header.nextElementSibling; // Get the next sibling (submenu)
  if (submenu.style.display === "block") {
    submenu.style.display = "none"; // Collapse
  } else {
    submenu.style.display = "block"; // Expand
  }
}
