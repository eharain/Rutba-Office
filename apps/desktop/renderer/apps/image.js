// Placeholder — replaced below in this build.
import React from 'react';
import { Empty } from '@rutba/office-ui';
import { AppFrame, useAppMenu } from '../shell.js';

export default function App({ app, shell }) {
  const menu = useAppMenu({ shell, appKey: 'image' });
  return (
    <AppFrame app={app} shell={shell} title={app.name} menu={menu} status={<span>{app.name}</span>}>
      <Empty icon="{image}" title={app.name}>Coming up in this build.</Empty>
    </AppFrame>
  );
}
