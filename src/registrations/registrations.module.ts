import { Module } from '@nestjs/common';
import { MongooseModule } from '@nestjs/mongoose';
import { RegistrationsController } from './registrations.controller';
import { RegistrationsService } from './registrations.service';
import {
  Registrations,
  RegistrationsSchema,
} from './schemas/registrations.schema';

@Module({
  imports: [
    MongooseModule.forFeature([
      {
        name: Registrations.name,
        schema: RegistrationsSchema,
      },
    ]),
  ],
  controllers: [RegistrationsController],
  providers: [RegistrationsService],
})
export class RegistrationsModule {}
