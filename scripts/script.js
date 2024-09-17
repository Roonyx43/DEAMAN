window.onload = function() {
    let count = 1;
    const maxCount = 10;

    // Adicionar novo conjunto de campos
    document.getElementById('add-btn').addEventListener('click', () => {
        if (count >= maxCount) {
            alert('O número máximo de 10 grupos de equipamentos foi atingido.');
            return;
        }

        count++;

        const inputContainer = document.getElementById('input-container');

        // Criar os novos elementos de input com IDs únicos
        const newFields = `
            <hr>
            <div class="container-equipamento">
                <label for="equipamento">Equipamento ${count}</label>
                <input type="text" id="eqp${count}" class="equipamento">
            </div>
            <div class="container-qntnovo">
                <label for="qntnovo">Novo</label>
                <input type="number" id="novo${count}" class="qntnovo">
            </div>
            <div class="container-revit">
                <label for="revit">Revitalizado</label>
                <input type="number" id="revi${count}" class="revit">
            </div>
            <div class="container-desc">
                <label for="desc">Descarte</label>
                <input type="number" id="desc${count}" class="desc">
            </div>
        `;

        inputContainer.insertAdjacentHTML('beforeend', newFields);
    });

    // Remover último conjunto de campos
    document.getElementById('remove-btn').addEventListener('click', () => {
        if (count > 1) {
            const inputContainer = document.getElementById('input-container');
            for (let i = 0; i < 5; i++) {
                inputContainer.removeChild(inputContainer.lastElementChild);
            }
            count--;
        }
    });
}
