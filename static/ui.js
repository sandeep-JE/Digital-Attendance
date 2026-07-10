const SIDEBAR_STORAGE_KEY = "facial-attendance-sidebar-collapsed";

function closeToast(toast) {
    if (!toast || toast.classList.contains("is-closing")) {
        return;
    }
    toast.classList.add("is-closing");
    window.setTimeout(() => toast.remove(), 180);
}

function initToasts() {
    document.querySelectorAll("[data-toast]").forEach((toast) => {
        toast.querySelector("[data-toast-close]")?.addEventListener("click", () => closeToast(toast));
        window.setTimeout(() => closeToast(toast), 10000);
    });
}

function initConfirmations() {
    document.querySelectorAll("form[data-confirm]").forEach((form) => {
        form.addEventListener("submit", (event) => {
            if (!window.confirm(form.dataset.confirm)) {
                event.preventDefault();
            }
        });
    });
}

function initSidebar() {
    const shell = document.querySelector("[data-app-shell]");
    const sidebar = document.querySelector("[data-sidebar]");
    const toggle = document.querySelector("[data-sidebar-toggle]");
    const closeButton = document.querySelector("[data-sidebar-close]");
    const backdrop = document.querySelector("[data-sidebar-backdrop]");
    if (!shell || !sidebar || !toggle) {
        return;
    }

    const desktopQuery = window.matchMedia("(min-width: 1024px)");

    function getStoredState() {
        try {
            return window.localStorage.getItem(SIDEBAR_STORAGE_KEY) === "true";
        } catch (error) {
            return false;
        }
    }

    function storeState(collapsed) {
        try {
            window.localStorage.setItem(SIDEBAR_STORAGE_KEY, String(collapsed));
        } catch (error) {
            // The layout still works when storage is disabled by the browser.
        }
    }

    function closeMobileSidebar({ restoreFocus = false } = {}) {
        document.body.classList.remove("sidebar-open");
        toggle.setAttribute("aria-expanded", "false");
        if (!desktopQuery.matches) {
            sidebar.setAttribute("aria-hidden", "true");
            sidebar.setAttribute("inert", "");
        }
        if (restoreFocus) {
            toggle.focus();
        }
    }

    function syncLayout() {
        if (desktopQuery.matches) {
            document.body.classList.remove("sidebar-open");
            sidebar.setAttribute("aria-hidden", "false");
            sidebar.removeAttribute("inert");
            shell.classList.toggle("sidebar-collapsed", getStoredState());
            toggle.setAttribute("aria-expanded", String(!shell.classList.contains("sidebar-collapsed")));
        } else {
            shell.classList.remove("sidebar-collapsed");
            closeMobileSidebar();
        }
    }

    toggle.addEventListener("click", () => {
        if (desktopQuery.matches) {
            const collapsed = shell.classList.toggle("sidebar-collapsed");
            storeState(collapsed);
            toggle.setAttribute("aria-expanded", String(!collapsed));
            return;
        }

        const willOpen = !document.body.classList.contains("sidebar-open");
        document.body.classList.toggle("sidebar-open", willOpen);
        toggle.setAttribute("aria-expanded", String(willOpen));
        if (willOpen) {
            sidebar.setAttribute("aria-hidden", "false");
            sidebar.removeAttribute("inert");
            closeButton?.focus();
        }
    });

    closeButton?.addEventListener("click", () => closeMobileSidebar({ restoreFocus: true }));
    backdrop?.addEventListener("click", () => closeMobileSidebar({ restoreFocus: true }));
    document.querySelectorAll(".sidebar-nav a").forEach((link) => {
        link.addEventListener("click", () => {
            if (!desktopQuery.matches) {
                closeMobileSidebar();
            }
        });
    });
    document.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && document.body.classList.contains("sidebar-open")) {
            closeMobileSidebar({ restoreFocus: true });
        }
    });
    desktopQuery.addEventListener("change", syncLayout);
    syncLayout();
}

function initUserMenu() {
    const menu = document.querySelector("[data-user-menu]");
    const toggle = menu?.querySelector("[data-user-menu-toggle]");
    const dropdown = menu?.querySelector("[data-user-dropdown]");
    if (!menu || !toggle || !dropdown) {
        return;
    }

    function closeMenu({ restoreFocus = false } = {}) {
        dropdown.hidden = true;
        toggle.setAttribute("aria-expanded", "false");
        if (restoreFocus) {
            toggle.focus();
        }
    }

    toggle.addEventListener("click", () => {
        const willOpen = dropdown.hidden;
        dropdown.hidden = !willOpen;
        toggle.setAttribute("aria-expanded", String(willOpen));
        if (willOpen) {
            dropdown.querySelector("[role='menuitem']")?.focus();
        }
    });
    document.addEventListener("click", (event) => {
        if (!menu.contains(event.target)) {
            closeMenu();
        }
    });
    menu.addEventListener("keydown", (event) => {
        if (event.key === "Escape" && !dropdown.hidden) {
            closeMenu({ restoreFocus: true });
        }
    });
}

document.addEventListener("DOMContentLoaded", () => {
    initToasts();
    initConfirmations();
    initSidebar();
    initUserMenu();
});
