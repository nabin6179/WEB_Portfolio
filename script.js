// The Digital Identity Project — script.js (simple version)

// Mobile menu toggle
const menuBtn = document.getElementById("menuBtn");
const navMenu = document.getElementById("navMenu");

menuBtn.addEventListener("click", function () {
  navMenu.classList.toggle("show");
});

// Close menu when a link is clicked (mobile)
const navLinks = document.querySelectorAll("#navMenu a");
navLinks.forEach(function (link) {
  link.addEventListener("click", function () {
    navMenu.classList.remove("show");
  });
});

// Dynamic project filter
const filterButtons = document.querySelectorAll(".filter-btn");
const projects = document.querySelectorAll(".project");

filterButtons.forEach(function (button) {
  button.addEventListener("click", function () {
    // highlight the active button
    filterButtons.forEach(function (btn) {
      btn.classList.remove("active");
    });
    button.classList.add("active");

    const filter = button.getAttribute("data-filter");

    // show or hide each project card
    projects.forEach(function (project) {
      const category = project.getAttribute("data-category");
      if (filter === "all" || category.includes(filter)) {
        project.classList.remove("hidden");
      } else {
        project.classList.add("hidden");
      }
    });
  });
});
