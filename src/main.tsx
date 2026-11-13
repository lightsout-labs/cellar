import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { webPlatform } from './app/platform/web';
import { App } from './app/ui/App';
import './app/ui/base.css';

declare const __COMMIT__: string;
document.documentElement.dataset.commit = __COMMIT__;

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App platform={webPlatform} />
  </StrictMode>,
);
