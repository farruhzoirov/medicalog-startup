import { DateTime } from 'luxon';

export const createDateRangeFilter = (fromDate?: string, toDate?: string) => {
  const filter: Record<string, Date> = {};

  if (fromDate) {
    const startDate = DateTime.fromISO(fromDate, { zone: 'Asia/Tashkent' })
      .startOf('day')
      .toUTC()
      .toJSDate();
    filter.$gte = startDate;
  }

  if (toDate) {
    const endDate = DateTime.fromISO(toDate, { zone: 'Asia/Tashkent' })
      .endOf('day')
      .toUTC()
      .toJSDate();
    filter.$lte = endDate;
  }

  return Object.keys(filter).length > 0 ? filter : null;
};
