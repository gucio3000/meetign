global.fetch = jest.fn();

test('findMeetingTimes API returns stubbed slots', async () => {
  const response = { slots: ['09:00'] };
  fetch.mockResolvedValue({ json: async () => response });

  const res = await fetch('/api/findMeetingTimes');
  const data = await res.json();

  expect(fetch).toHaveBeenCalledWith('/api/findMeetingTimes');
  expect(data).toEqual(response);
});
