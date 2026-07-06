<script>
function safeTrack(eventName, params) {
  if (typeof analyticsTrack === "function") {
    analyticsTrack(eventName, params || {});
  }
}

document.addEventListener("DOMContentLoaded", function () {
  var burger = document.getElementById("ozmBurger");
  var nav = document.getElementById("ozmNav");
  var dropdownButtons = document.querySelectorAll(".ozm-has-sub > .ozm-nav-btn");

  if (burger && nav) {
    burger.addEventListener("click", function () {
      var isOpen = nav.classList.toggle("is-open");

      burger.classList.toggle("is-open", isOpen);
      burger.setAttribute("aria-expanded", isOpen ? "true" : "false");

      if (!isOpen) {
        closeAllDropdowns();
      }
    });
  }

  dropdownButtons.forEach(function (btn) {
    btn.addEventListener("click", function () {
      if (window.innerWidth > 1024) {
        return;
      }

      var item = btn.parentElement;
      var isOpen = item.classList.contains("is-open");

      closeAllDropdowns();

      if (!isOpen) {
        item.classList.add("is-open");
      }
    });
  });

  document.addEventListener("click", function (e) {
    if (window.innerWidth <= 1024) {
      return;
    }

    if (!e.target.closest(".ozm-has-sub")) {
      closeAllDropdowns();
    }
  });

  window.addEventListener("resize", function () {
    if (window.innerWidth > 1024) {
      if (nav) nav.classList.remove("is-open");
      if (burger) {
        burger.classList.remove("is-open");
        burger.setAttribute("aria-expanded", "false");
      }
      closeAllDropdowns();
    }
  });

  function closeAllDropdowns() {
    var openItems = document.querySelectorAll(".ozm-has-sub.is-open");

    openItems.forEach(function (item) {
      item.classList.remove("is-open");
    });
  }
});
</script>