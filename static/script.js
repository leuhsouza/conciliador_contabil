document.addEventListener('DOMContentLoaded', function() {
    // Lógica para mostrar ou esconder opções baseadas no tipo selecionado
    const tipoElement = document.getElementById('tipo');
    if (tipoElement) {
        tipoElement.addEventListener('change', function() {
            const opcaoRazao = document.getElementById('opcaoRazao');
            if (this.value === 'Razao') {
                opcaoRazao.classList.remove('hidden');
            } else {
                opcaoRazao.classList.add('hidden');
            }
        });
    }

    // Função para retornar à página inicial
    function voltarParaInicio() {
        window.location.href = '/';
    }

    // Lógica de soma e destaque de linhas
    let totalSum = 0;
    let selectedIds = []; // Lista para armazenar os IDs das linhas destacadas

    function updateSum(row, isAdding) {
        const dValue = parseFloat(row.querySelector('td:nth-child(7)').textContent.replace(',', '.')) || 0; // Coluna D
        const cValue = parseFloat(row.querySelector('td:nth-child(8)').textContent.replace(',', '.')) || 0; // Coluna C

        if (isAdding) {
            totalSum += dValue;
            totalSum -= cValue;
        } else {
            totalSum -= dValue;
            totalSum += cValue;
        }

        document.getElementById('sumValue').textContent = totalSum.toFixed(2);
    }

    let selectedRow = null;

    document.querySelectorAll('table.data tr').forEach(function(row) {
        row.addEventListener('click', function() {
            if (selectedRow) {
                selectedRow.classList.remove('selected');
            }
            selectedRow = this;
            this.classList.add('selected');
            this.focus();
        });

        row.addEventListener('dblclick', function() {
            const rowId = this.querySelector('td:first-child').textContent; // Supondo que o ID esteja na primeira coluna
            if (this.classList.contains('highlight')) {
                updateSum(this, false);
                this.classList.remove('highlight');
                selectedIds = selectedIds.filter(id => id !== rowId); // Remove o ID da lista
            } else {
                updateSum(this, true);
                this.classList.add('highlight');
                selectedIds.push(rowId); // Adiciona o ID à lista
            }

            // Atualiza o campo oculto com os IDs das linhas destacadas
            const selectedIdsInput = document.getElementById('selectedIds');
            if (selectedIdsInput) {
                selectedIdsInput.value = selectedIds.join(',');
            }
        });
    });

    document.addEventListener('keydown', function(event) {
        if (!selectedRow) return;

        const rows = Array.from(document.querySelectorAll('table.data tr'));
        const index = rows.indexOf(selectedRow);

        if (event.key === 'ArrowDown') {
            event.preventDefault();
            if (index < rows.length - 1) {
                selectedRow.classList.remove('selected');
                selectedRow = rows[index + 1];
                selectedRow.classList.add('selected');
                selectedRow.focus();
            }
        } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            if (index > 0) {
                selectedRow.classList.remove('selected');
                selectedRow = rows[index - 1];
                selectedRow.classList.add('selected');
                selectedRow.focus();
            }
        } else if (event.key === 'Enter') {
            event.preventDefault();
            selectedRow.dispatchEvent(new Event('dblclick'));
        }
    });

    // Função para salvar a conciliação
    const saveConciliationBtn = document.getElementById('saveConciliation');
    if (saveConciliationBtn) {
        saveConciliationBtn.addEventListener('click', function() {
            let selectedIds = [];
            document.querySelectorAll('table.data tr.highlight').forEach(function(row) {
                const id = row.querySelector('td:first-child').textContent; // Supondo que a primeira coluna seja o ID
                if (id) {
                    selectedIds.push(id.trim());
                }
            });

            if (selectedIds.length > 0) {
                fetch('/save_conciliation', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({ ids: selectedIds })
                })
                .then(response => response.text())
                .then(() => {
                    alert('Conciliação salva com sucesso!');
                })
                .catch(() => {
                    alert('Erro ao salvar a conciliação.');
                });
            } else {
                alert('Nenhuma linha selecionada para conciliação.');
            }
        });
    }
});
