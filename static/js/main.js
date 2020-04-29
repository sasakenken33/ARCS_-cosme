document.addEventListener('DOMContentLoaded', function () {
  document.querySelector("#btn").addEventListener("click", function() {
    let status = document.querySelector(".loader.hidden")
    status.classList.remove("hidden");
  }, false);
}, false);