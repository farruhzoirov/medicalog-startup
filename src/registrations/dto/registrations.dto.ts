import {
  IsDateString,
  IsEnum,
  IsMongoId,
  IsNotEmpty,
  IsNumber,
  IsOptional,
  IsString,
  Length,
} from 'class-validator';
import { Gender } from '../schemas/registrations.schema';

export class AuthDto {
  @IsString()
  username: string;

  @IsString()
  password: string;
}


export class ForGenerateWordDto {
  @IsOptional()
  @IsDateString()
  createdAtFrom?: string;

  @IsOptional()
  @IsDateString()
  createdAtTo?: string;
}

export class GetFilteredRegistrationsDto {
  @IsOptional()
  @IsNumber()
  page: number;

  @IsOptional()
  @IsNumber()
  limit: number;

  @IsOptional()
  @IsString()
  search?: string;

  @IsOptional()
  @IsString()
  birthDateFrom?: string;

  @IsOptional()
  @IsString()
  birthDateTo?: string;

  @IsOptional()
  @IsNumber()
  ageFrom?: string;

  @IsOptional()
  @IsNumber()
  ageTo?: string;

  @IsOptional()
  @IsDateString()
  createdAtFrom?: string;

  @IsOptional()
  @IsDateString()
  createdAtTo?: string;

  @IsOptional()
  @IsString()
  address?: string;

  @IsOptional()
  @IsString()
  otherAddress?: string;

  @IsOptional()
  @IsString()
  job?: string;

  @IsOptional()
  @IsString()
  otherJob?: string;

  @IsOptional()
  @IsString()
  visitReason?: string;

  @IsOptional()
  @IsString()
  otherVisitReason?: string;

  @IsOptional()
  @IsEnum(Gender)
  gender?: string;

  @IsOptional()
  @IsString()
  radiologyReport: string;

  @IsOptional()
  @IsString()
  otherRadiologyReport: string;
}

export class CreateRegistrationDto {
  @IsString()
  @IsNotEmpty()
  fullName: string;

  @IsString()
  @IsOptional()
  address?: string;

  @IsString()
  @IsOptional()
  otherAddress?: string;

  @IsString()
  @Length(4, 4)
  birthDate: string;

  @IsEnum(Gender)
  @IsNotEmpty()
  gender: Gender;

  @IsString()
  @IsOptional()
  job?: string;

  @IsString()
  @IsOptional()
  otherJob?: string;

  @IsString()
  @IsOptional()
  visitReason?: string;

  @IsString()
  @IsOptional()
  otherVisitReason?: string;

  @IsString()
  @IsNotEmpty()
  radiationDose: string;

  @IsString()
  @IsOptional()
  radiologyReport?: string;

  @IsString()
  @IsOptional()
  otherRadiologyReport?: string;

  @IsString()
  @IsOptional()
  phone?: string;
}

export class UpdateRegistrationDto {
  @IsMongoId()
  @IsString()
  id: string;

  @IsString()
  @IsOptional()
  fullName?: string;

  @IsString()
  @IsOptional()
  address?: string;

  @IsString()
  @IsOptional()
  otherAddress?: string;

  @IsString()
  @Length(4, 4)
  @IsOptional()
  birthDate?: string;

  @IsEnum(Gender)
  @IsOptional()
  gender?: Gender;

  @IsString()
  @IsOptional()
  job?: string;

  @IsString()
  @IsOptional()
  otherJob?: string;

  @IsString()
  @IsOptional()
  visitReason?: string;

  @IsString()
  @IsOptional()
  otherVisitReason?: string;

  @IsString()
  @IsOptional()
  radiationDose?: string;

  @IsString()
  @IsOptional()
  radiologyReport?: string;

  @IsString()
  @IsOptional()
  otherRadiologyReport?: string;

  @IsString()
  @IsOptional()
  phone?: string;
}

export class DeleteRegistrationDto {
  @IsMongoId()
  @IsString()
  id: string;
}
