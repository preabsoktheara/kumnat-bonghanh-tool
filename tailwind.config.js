/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      fontFamily: {
        siemreap: ['"Khmer OS Siemreap"', 'Siemreap', 'sans-serif'],
        moul: ['"Khmer M1"', '"Khmer OS Muol Light"', 'Moul', 'serif'],
      },
    },
  },
  plugins: [],
};
