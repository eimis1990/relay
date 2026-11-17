import type { Metadata } from 'next';
import './globals.css';
export const metadata:Metadata={title:'Relay — Make room for the work that matters',description:'Managed email operations. A clearer inbox, an organized work list, and a human in control.'};
export default function Layout({children}:{children:React.ReactNode}){return <html lang="en"><body>{children}</body></html>}
