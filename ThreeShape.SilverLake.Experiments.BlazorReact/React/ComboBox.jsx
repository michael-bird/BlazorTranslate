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
        onChange={(e) => onChange(props.dotNetHelper, e)}
    />
);

const onChange = (dotNetHelper, e) => {
    dotNetHelper.invokeMethodAsync('SetSelected', e.target.value.Name);
};

export default ComboBox;