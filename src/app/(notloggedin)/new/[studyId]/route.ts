import {prisma} from '~/server/db';
import {createNewParticipation} from '~/server/api/routers/participation';
import {redirect} from 'next/navigation';

export async function GET(request: Request, {params}: {params: Promise<{studyId: string}>}) {
  const {studyId} = await params;
  const url = new URL(request.url);
  const rid = url.searchParams.get('rid') ?? undefined;
  console.log('rid from URL:', rid);  

  const study = await prisma.study.findUnique({
    where: {
      id: studyId,
    },
  });

  if (!study) {
    redirect('/');
  }

  if (!study.openParticipation) {
    redirect('/');
  }

  const {code} = await createNewParticipation(prisma, study.id, rid);

  redirect('/' + code);
}
