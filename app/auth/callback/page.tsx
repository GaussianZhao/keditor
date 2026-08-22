'use client';

import { broadcastResponseToMainFrame } from '@azure/msal-browser/redirect-bridge';
import { useEffect, useState } from 'react';

export default function MicrosoftAuthCallback() {
  const [error, setError] = useState('');

  useEffect(() => {
    broadcastResponseToMainFrame().catch((reason: unknown) => {
      setError(reason instanceof Error ? reason.message : 'Microsoft sign-in could not be completed.');
    });
  }, []);

  return (
    <main style={{ minHeight: '100vh', display: 'grid', placeItems: 'center', padding: 24 }}>
      <p>{error || 'Connecting to OneDrive…'}</p>
    </main>
  );
}
