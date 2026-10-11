alter table public.documents
    drop constraint if exists documents_ext_check;
alter table public.documents
    add constraint documents_ext_check
    check (ext in ('pdf', 'pptx', 'docx', 'txt', 'mp4', 'mov', 'webm', 'opus', 'ogg', 'mp3', 'm4a', 'wav', 'aac'));

update storage.buckets
   set allowed_mime_types = array[
       'application/pdf',
       'application/vnd.openxmlformats-officedocument.presentationml.presentation',
       'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
       'text/plain',
       'application/octet-stream',
       'video/mp4',
       'video/quicktime',
       'video/webm',
       'audio/ogg',
       'audio/mpeg',
       'audio/mp4',
       'audio/wav',
       'audio/aac'
   ]
 where id = 'study-materials';
