import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'app.paperline.office',
  appName: 'Paperline',
  webDir: 'dist',
  backgroundColor: '#f4f5fa',
  android: { allowMixedContent: false },
  ios: { contentInset: 'never' },
};

export default config;
