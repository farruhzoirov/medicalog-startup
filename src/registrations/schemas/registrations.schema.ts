import { Prop, Schema, SchemaFactory } from '@nestjs/mongoose';
export enum Gender {
  MALE = 'male',
  FEMALE = 'female',
}

export type RegistrationsDocument = Registrations & Document;

@Schema({ timestamps: true })
export class Registrations {
  @Prop({ type: Number, required: true, default: 1 })
  yearlyCount: number;

  @Prop({ type: Number, required: true, default: 1 })
  radiologyFilmNumber: number;

  @Prop({ type: Number, required: true, default: 1 })
  dailyCount: number;

  @Prop({ required: true })
  fullName: string;

  @Prop({ default: null })
  address: string;

  @Prop({ default: null })
  otherAddress: string;

  @Prop({ required: true })
  birthDate: string;

  @Prop({ required: true })
  age: number;

  @Prop({ required: true, enum: Gender })
  gender: Gender;

  @Prop({ default: null })
  job: string;

  @Prop({ default: null })
  otherJob: string;

  @Prop({ default: null })
  visitReason: string;

  @Prop({ default: null })
  otherVisitReason: string;

  @Prop({ required: true })
  radiationDose: string;

  @Prop({ default: null })
  radiologyReport: string;

  @Prop({ default: null })
  otherRadiologyReport: string;

  @Prop({ default: null })
  phone: string;
}

export const RegistrationsSchema = SchemaFactory.createForClass(Registrations);
// Compound index
RegistrationsSchema.index({
  fullName: 1,
  phone: 1,
  adress: 1,
  otherAddress: 1,
  job: 1,
  otherJob: 1,
  visitReason: 1,
  otherVisitReason: 1,
  radiologyReport: 1,
  otherRadioLogyReport: 1,
});
RegistrationsSchema.index({ visitReason: 1, otherVisitReason: 1 });
RegistrationsSchema.index({ job: 1, otherJob: 1 });

// Single index
RegistrationsSchema.index({ birthDate: 1 });
RegistrationsSchema.index({ createdAt: 1 });
RegistrationsSchema.index({ gender: 1 });
