import { jobList } from '../registrations/data/job-list.constant';
import { mahallaList } from '../registrations/data/mahalla-list.constant';
import { reportList } from '../registrations/data/report-list.constant';
import { visitReasons } from '../registrations/data/visit-reasons.constant';

export interface RegistrationData {
    _id: string;
    yearlyCount: number;
    radiologyFilmNumber: number;
    dailyCount: number;
    fullName: string;
    address: string;
    otherAddress?: string;
    birthDate: string;
    age: number;
    gender: string;
    job: string;
    otherJob?: string;
    visitReason: string;
    otherVisitReason?: string;
    radiationDose: string;
    radiologyReport: string;
    otherRadiologyReport?: string;
    phone: string;
    createdAt: string;
    updatedAt: string;
}

export interface FormattedRegistrationData extends Omit<RegistrationData, 'address' | 'job' | 'visitReason' | 'radiologyReport'> {
    address: string;
    job: string;
    visitReason: string;
    radiologyReport: string;
    gender: string;
}

/**
 * @param registrations
 * @returns
 */


export function formatRegistrations(registrations: RegistrationData[]): FormattedRegistrationData[] {
    return registrations.map(registration => {
        const address = registration.address === 'other' ? registration.otherAddress : mahallaList[registration.address];
        const job = registration.job === 'other' ? registration.otherJob : jobList[registration.job];
        const visitReason = registration.visitReason === 'other' ? registration.otherVisitReason : visitReasons[registration.visitReason];
        const radiologyReport = registration.radiologyReport === 'other' ? registration.otherRadiologyReport : reportList[registration.radiologyReport];
        const gender = registration.gender === 'male' ? 'Erkak' : 'Ayol';
        return {
            ...registration,
            address,
            job,
            visitReason,
            radiologyReport,
            gender
        }
    });
}

/**
 * Bitta registration ma'lumotini formatlash
 * @param registration - Bitta registration ma'lumoti
 * @returns Formatlangan registration ma'lumoti
 */
export function formatSingleRegistration(registration: RegistrationData): FormattedRegistrationData {
    return formatRegistrations([registration])[0];
}

/**
 * Key-larni name-larga o'tkazish uchun umumiy helper funksiya
 * @param key - Qidirilayotgan key
 * @param dataObject - Ma'lumotlar obyekti
 * @returns Name yoki key (agar topilmasa)
 */
export function getDisplayName(key: string, dataObject: Record<string, string>): string {
    return dataObject[key] || key;
}

/**
 * Barcha constant-larni bitta obyektda qaytarish
 */
export const registrationConstants = {
    visitReasons,
    jobList,
    mahallaList,
    reportList,
};
