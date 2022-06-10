window.addEventListener('load', () => {
    const form = document.querySelector('#new-order-form');
    const input_name = document.querySelector('#new-order-name');
    const input_price = document.querySelector('#new-order-price');
    const list_el = document.querySelector('.list');    

    form.addEventListener('submit', (e) => {
        e.preventDefault();

        const new_name = input_name.value;
        const new_price = input_price.value;

        if(!new_name || !new_price){
            alert("警告 請輸入完整資料!");
            return;
        }

        const order_el = document.createElement('div');
        order_el.classList.add('content');  // the head

        const order_name_el = document.createElement('input');
        order_name_el.classList.add('text');
        order_name_el.type = 'text';
        order_name_el.value = new_name;
        order_name_el.setAttribute('readonly','readonly');
        
        order_el.appendChild(order_name_el);

        const order_price_el = document.createElement('input');
        order_price_el.classList.add('text');
        order_price_el.type = 'text';
        order_price_el.setAttribute('value', new_price);
        order_price_el.setAttribute('readonly','readonly');

        order_el.appendChild(order_price_el);
        
        let order_amount_el = document.createElement('input');
        order_amount_el.classList.add('text');
        order_amount_el.type = 'text';
        order_amount_el.setAttribute('value', '0');
        order_amount_el.setAttribute('readonly','readonly'); 
        
        order_el.appendChild(order_amount_el);
        
        let order_total_el = document.createElement('input');
        order_total_el.classList.add('text');
        order_total_el.type = 'text';
        order_total_el.setAttribute('value', '0');
        order_total_el.setAttribute('readonly','readonly'); 
        
        order_el.appendChild(order_total_el);
        
        const order_action = document.createElement('div');
        order_action.classList.add('action');
        
        const button_add = document.createElement('button');
        button_add.classList.add('add');
        button_add.innerHTML = '+1'
        
        const button_sub = document.createElement('button');
        button_sub.classList.add('sub');
        button_sub.innerHTML = '-1'
        
        const button_delete = document.createElement('button');
        button_delete.classList.add('delete');
        button_delete.innerHTML = '刪除'
        
        order_action.appendChild(button_add);
        order_action.appendChild(button_sub);
        order_action.appendChild(button_delete);

        order_el.appendChild(order_action);

        list_el.appendChild(order_el);

        input_name.value = '';
        input_price.value = '';

        button_add.addEventListener('click', () => {
            let amount = parseInt(order_amount_el.value) + 1;
            let total = parseInt(order_total_el.value) + parseInt(order_price_el.value);
            
            order_amount_el.value = amount;
            order_total_el.value = total;
        })
        
        button_sub.addEventListener('click', () => {
            if(order_amount_el.value > 0){
                let amount = parseInt(order_amount_el.value) - 1;
                let total = parseInt(order_total_el.value) - parseInt(order_price_el.value);
    
                order_amount_el.value = amount;
                order_total_el.value = total;
            }
        })
        
        button_delete.addEventListener('click', () => {
            list_el.removeChild(order_el);
        })

        
    })
})