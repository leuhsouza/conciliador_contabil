document.addEventListener('DOMContentLoaded', function() {
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
});
function saveConciliation() {
    let selectedIds = [];
    let allIds = [];

    // Captura os IDs de todas as linhas exibidas na tabela
    document.querySelectorAll('table.data tr').forEach(function(row) {
        const idElement = row.querySelector('td:first-child');
        if (idElement) { // Verifica se idElement não é null
            const id = idElement.textContent;
            if (id) {
                allIds.push(id.trim());
                if (row.classList.contains('highlight')) {
                    selectedIds.push(id.trim());
                }
            }
        }
    });

    if (allIds.length > 0) {
        fetch('/save_conciliation', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ ids: selectedIds, all_ids: allIds })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                alert('Conciliação salva com sucesso!');
                window.location.reload(); // Atualiza a página para mostrar as alterações
            } else {
                alert('Erro ao salvar a conciliação: ' + data.message);
            }
        })
        .catch(() => {
            alert('Erro ao salvar a conciliação.');
        });
    } else {
        alert('Nenhuma linha encontrada para salvar.');
    }
}

// Função para remover a conciliação das linhas filtradas (apenas visual)
function removeConciliation() {
    const filterField = document.getElementById('filter_field').value.trim();
    const filterConta = document.getElementById('filter_conta').value.trim();

    document.querySelectorAll('table.data tr').forEach(row => {
        const historico = row.querySelector('td:nth-child(3)')?.textContent || '';
        const conta = row.querySelector('td:nth-child(10)')?.textContent || ''; // Ajustar conforme o índice da coluna

        // Verifica se a linha corresponde aos filtros aplicados
        if ((filterField === '' || historico.includes(filterField)) &&
            (filterConta === '' || conta.includes(filterConta))) {
            row.classList.remove('highlight');
        }
    });

    alert('Conciliação removida visualmente. Para salvar as alterações, clique em "Salvar Conciliação".');
}
