import React from 'react';
import ReactDOM from 'react-dom'
import ComboBox from './ComboBox'

const toothStuff = [
    {
        text: "Crown",
        id: 1,
    },
    {
        text: "Bridge",
        id: 2,
    },
    {
        text: "Additional",
        id: 3,
    },
    {
        text: "Implant",
        id: 4,
    },
    {
        text: "Coping",
        id: 5,
    },
];

export default function renderHello(element) {
    ReactDOM.render(<ComboBox data={toothStuff} textField="text" />, element);
}

window.renderReact = renderHello;