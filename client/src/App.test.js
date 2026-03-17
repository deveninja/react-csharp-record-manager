import { render, screen } from '@testing-library/react';
import App from './App';

beforeEach(() => {
  global.fetch = jest.fn(() =>
    Promise.resolve({
      ok: true,
      json: () => Promise.resolve([]),
    })
  );
});

test('renders record manager title', async () => {
  render(<App />);
  const heading = await screen.findByText(/record manager/i);
  expect(heading).toBeInTheDocument();
});
