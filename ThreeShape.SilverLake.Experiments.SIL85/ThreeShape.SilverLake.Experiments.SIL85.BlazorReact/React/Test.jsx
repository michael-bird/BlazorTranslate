import React from 'react';
import ReactDOM from 'react-dom'
import ComboBox from './ComboBox'

export default function renderHello(dotNetHelper, element, serialisedItems) {

    const items = JSON.parse(serialisedItems);

    ReactDOM.render(<ComboBox data={items} textField="Name" dotNetHelper={dotNetHelper} />, element);
}

window.renderReact = renderHello;