const body = document.body;
const toggle = document.getElementById("themeToggle");
const savedTheme = localStorage.getItem("theme");

if (savedTheme === "light") {
    body.classList.remove("theme-dark");
    body.classList.add("theme-light");
}

if (toggle) {
    toggle.addEventListener("click", () => {
        const isLight = body.classList.contains("theme-light");
        if (isLight) {
            body.classList.remove("theme-light");
            body.classList.add("theme-dark");
            localStorage.setItem("theme", "dark");
        } else {
            body.classList.remove("theme-dark");
            body.classList.add("theme-light");
            localStorage.setItem("theme", "light");
        }
    });
}

const splitMode = document.getElementById("splitMode");
const rangeBox = document.getElementById("rangeBox");

if (splitMode && rangeBox) {
    const syncRangeVisibility = () => {
        rangeBox.style.display = splitMode.value === "range" ? "block" : "none";
    };
    splitMode.addEventListener("change", syncRangeVisibility);
    syncRangeVisibility();
}
