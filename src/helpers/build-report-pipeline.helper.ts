import type { PipelineStage } from "mongoose";

/**
 * Builds aggregation pipeline for patient report stats.
 * If from/to are not provided, no $match is applied (all data) and dateFrom/dateTo come from min/max createdAt.
 */
export const buildReportPipeline = (from?: Date, to?: Date): PipelineStage[] => {
  const stages: PipelineStage[] = [];

  if (from != null && to != null) {
    stages.push({
      $match: {
        createdAt: {
          $gte: from,
          $lte: to,
        },
      },
    });
  }

  stages.push({
    $group: {
      _id: null,
      total: { $sum: 1 },
      women: {
        $sum: { $cond: [{ $eq: ["$gender", "female"] }, 1, 0] },
      },
      men: {
        $sum: { $cond: [{ $eq: ["$gender", "male"] }, 1, 0] },
      },
      unemployed: {
        $sum: {
          $cond: [
            { $or: [{ $eq: ["$job", "unemployed"] }, { $eq: ["$otherJob", "Ishsiz"] }] },
            1,
            0,
          ],
        },
      },
      pensioners: {
        $sum: {
          $cond: [
            { $or: [{ $eq: ["$job", "pensioner"] }, { $regexMatch: { input: { $ifNull: ["$otherJob", ""] }, regex: /nafaqaxo'r|pensioner/i } }] },
            1,
            0,
          ],
        },
      },
      disabled: {
        $sum: {
          $cond: [
            {
              $or: [
                { $eq: ["$job", "disabled"] },
                { $regexMatch: { input: { $ifNull: ["$otherJob", ""] }, regex: /nogiron|imkoniyati cheklangan|disabled/i } },
              ],
            },
            1,
            0,
          ],
        },
      },
      dateFrom: { $min: "$createdAt" },
      dateTo: { $max: "$createdAt" },
    },
  });

  stages.push({
    $project: {
      _id: 0,
      total: 1,
      women: 1,
      men: 1,
      unemployed: 1,
      pensioners: 1,
      disabled: 1,
      dateFrom: 1,
      dateTo: 1,
    },
  });

  return stages;
};
