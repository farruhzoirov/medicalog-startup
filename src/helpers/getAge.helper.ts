import { BadRequestException } from '@nestjs/common';

export async function getAgeHelper(dateStr: string): Promise<number> {
  const today = new Date();
  let yearDifference = today.getFullYear() - parseInt(dateStr);
  if (yearDifference < 0) {
    throw new BadRequestException('Year is invalid');
  }

  return yearDifference;
}
