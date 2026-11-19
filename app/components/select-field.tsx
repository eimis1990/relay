'use client';
import * as Select from '@radix-ui/react-select';
import {Check,ChevronDown,ChevronUp} from 'lucide-react';

type Option=string|{value:string;label:string;color?:string};
export default function SelectField({label,value,onChange,options,disabled=false,id}:{label:string;value:string;onChange:(value:string)=>void;options:readonly Option[];disabled?:boolean;id?:string}){
 return <Select.Root value={value} onValueChange={onChange} disabled={disabled}>
  <Select.Trigger id={id} className="r-select-trigger" aria-label={label}>
   <Select.Value/><Select.Icon className="r-select-chevron"><ChevronDown size={17}/></Select.Icon>
  </Select.Trigger>
  <Select.Portal><Select.Content className="r-select-menu" position="popper" sideOffset={7} collisionPadding={12}>
   <Select.ScrollUpButton className="r-select-scroll"><ChevronUp size={16}/></Select.ScrollUpButton>
   <Select.Viewport className="r-select-viewport">{options.map(option=>{const o=typeof option==='string'?{value:option,label:option}:option;return <Select.Item key={o.value} value={o.value} textValue={o.label} className="r-select-option">
    {o.color&&<span aria-hidden="true" className={'r-color-dot '+o.color}/>}
    <Select.ItemText>{o.label}</Select.ItemText><Select.ItemIndicator className="r-select-check"><Check size={16}/></Select.ItemIndicator>
   </Select.Item>})}</Select.Viewport>
   <Select.ScrollDownButton className="r-select-scroll"><ChevronDown size={16}/></Select.ScrollDownButton>
  </Select.Content></Select.Portal>
 </Select.Root>;
}
