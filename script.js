(() => {
    const removeEyeButton = () => {
        document.querySelectorAll('button[aria-label="Show password text"]').forEach(btn => btn.remove());
    };
    const observer = new MutationObserver(removeEyeButton);
    observer.observe(document.body, { childList: true, subtree: true });
    removeEyeButton();
})();