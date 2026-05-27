/* Главная страница списка шаблонов */

document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('.folder-header').forEach(function (header) {
        header.addEventListener('click', function () {
            var content = this.nextElementSibling;
            var isClosed = content.style.display === 'none' || content.style.display === '';
            content.style.display = isClosed ? 'block' : 'none';
            this.classList.toggle('open', isClosed);
        });
    });
});
