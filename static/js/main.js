document.addEventListener('DOMContentLoaded', function () {
  document.querySelector("#btn").addEventListener("click", function() {
    let loading_btn = document.querySelector("#btn");
    loading_btn.value = "レビュー収集中"
    let loading_title = document.querySelector("h2.loading");
    loading_title.textContent = "レビュー収集中";
  }, false);
}, false);