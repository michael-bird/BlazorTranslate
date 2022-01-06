import React from 'react';
import { ComboBox as CommonComboBox } from '@progress/kendo-react-dropdowns';
import '@progress/kendo-theme-material/dist/all.css';
import './kendo-material.scss'
import './DropDown.scss';

const ComboBox = props => (
    <CommonComboBox
        {...props}
        ref={elem => (props.setRef ? props.setRef(elem) : {})}
        popupSettings={{
            className: 'dropdown-popup',
        }}
    />
);

export default ComboBox;