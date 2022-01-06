import React from 'react';
import ReactDOM from 'react-dom'
import ComboBox from './ComboBox'

export default function renderHello(element, serialisedItems) {
    console.log(serialisedItems);
    const items = JSON.parse(serialisedItems);
    console.log(items);
    ReactDOM.render(<ComboBox data={items} textField="Name" />, element);
}

window.renderReact = renderHello;