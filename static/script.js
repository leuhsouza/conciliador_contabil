document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('tipo').addEventListener('change', function() {
        var opcaoRazao = document.getElementById('opcaoRazao');
        if (this.value === 'Razao') {
            opcaoRazao.classList.remove('hidden');
        } else {
            opcaoRazao.classList.add('hidden');
        }
    });
});