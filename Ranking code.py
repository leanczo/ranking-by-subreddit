import praw
import sys
import xlsxwriter 
from time import time

reddit = praw.Reddit(client_id='XXXXXXXXX',
                     client_secret='XXXXXXXX', password='XXXXXXX',
                     user_agent='XXXXXX', username='XXXXXX')

non_bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)
InitialTime = time()
subreddit = reddit.subreddit('XXXXXXXX')                
i=0
ArrayCommnent=[]
ArrayPost=[]
ArrayFlair=[]
WorstCommentUpvote=99999
WorstComment=''
WorstAuthor=''
WorstCommentPost=''
BestCommentUpvote=-99999
BestComment=''
BestAuthor=''
BestCommentPost=''
PostQuantity=0
CommentQuantity=0
TotalPostUpvotes=0
TotalPostDownvotes=0
TotalCommentUpvotes=0
new_python = subreddit.new(limit=1500)

for submission in new_python:
    try:
        if int('1551398400')< int(submission.created_utc)<int('1551916800'):
            ratio = submission.upvote_ratio
            ups = round((ratio*submission.score)/(2*ratio - 1)) if ratio != 0.5 else round(submission.score/2)
            downs = ups - submission.score
            TotalPostUpvotes+=submission.ups
            TotalPostDownvotes+=downs
            UserPostQuantity=0
            PostQuantity+=1

            if '{}'.format(submission.author) not in ArrayPost:
                ArrayPost.insert(i,'{}'.format(submission.author))
                ArrayPost.insert(i+1,submission.ups-1)
                ArrayPost.insert(i+2,downs)
                ArrayPost.insert(i+3,UserPostQuantity+1)
                i+=4
            else:
                index=ArrayPost.index('{}'.format(submission.author))
                SumQuantityUpvotePost=int(ArrayPost[index+1])+ int(submission.ups-1)
                ArrayPost.insert(index+1,str(SumQuantityUpvotePost))
                del ArrayPost[index+2]
                SumQuantityDownvotePost=int(ArrayPost[index+2])+ downs
                ArrayPost.insert(index+2,str(SumQuantityDownvotePost))
                del ArrayPost[index+3]
                SumQuantityPostUser=int(ArrayPost[index+3])+ 1
                ArrayPost.insert(index+3,str(SumQuantityPostUser))
                del ArrayPost[index+4]
                i+=4

            if '{}'.format(submission.link_flair_text) not in ArrayFlair:
                ArrayFlair.insert(i,'{}'.format(submission.link_flair_text))
                ArrayFlair.insert(i+1,1)
                i+=2
            else:
                index=ArrayFlair.index('{}'.format(submission.link_flair_text))
                SumQuantityFlair=int(ArrayFlair[index+1])+ 1
                ArrayFlair.insert(index+1,str(SumQuantityFlair))
                del ArrayFlair[index+2]
                i+=2

            UserCommentQuantity=0
            submission.comments.replace_more(limit=0)

            for comment in submission.comments.list():
                TotalCommentUpvotes+=comment.ups
                CommentQuantity+=1

                if WorstCommentUpvote>comment.ups:
                    WorstComment=comment.body
                    WorstCommentUpvote=comment.ups
                    WorstAuthor='{}'.format(comment.author)
                    WorstCommentPost='{}'.format(submission.title)

                if BestCommentUpvote<comment.ups:
                    BestComment=comment.body
                    BestCommentUpvote=comment.ups
                    BestAuthor='{}'.format(comment.author)
                    BestCommentPost='{}'.format(submission.title)
  
                if '{}'.format(comment.author) not in ArrayCommnent:
                    ArrayCommnent.insert(i,'{}'.format(comment.author))
                    ArrayCommnent.insert(i+1,comment.ups-1)
                    ArrayCommnent.insert(i+2,UserCommentQuantity+1)
                    i+=3
                else:
                    index=ArrayCommnent.index('{}'.format(comment.author))
                    suma=int(ArrayCommnent[index+1])+ int(comment.ups-1)
                    ArrayCommnent.insert(index+1,str(suma))
                    del ArrayCommnent[index+2]
                    SumQuantityCommentUser=int(ArrayCommnent[index+2])+ 1
                    ArrayCommnent.insert(index+2,str(SumQuantityCommentUser))
                    del ArrayCommnent[index+3]
                    i+=3          
    except UnicodeEncodeError:
        print('Unicode Encode Error')
i=0 
print('Posts; Karma Total; Downvotes Posts; Quantity Posts')
for i in range(0, len(ArrayPost), 4):
    print(str(ArrayPost[i]), ';', str(ArrayPost[i+1]), ';', str(ArrayPost[i+2]), ';', str(ArrayPost[i+3]))
print()

i=0
print('Comments; Karma; Quantity Comments')
for i in range(0, len(ArrayCommnent), 3):
    print(str(ArrayCommnent[i]), ';', str(ArrayCommnent[i+1]), ';', str(ArrayCommnent[i+2]))
print()

i=0
print('Flairs; Quantity')
for i in range(0, len(ArrayFlair), 2):
    print(str(ArrayFlair[i].translate(non_bmp_map)), ';', str(ArrayFlair[i+1]))
print()

workbook = xlsxwriter.Workbook('Posts.xlsx') 
worksheet = workbook.add_worksheet() 
row = 0
column = 0

for item in ArrayPost: 
    worksheet.write(row, column, item)
    column += 1
    if (column  % 4) == 0:
        column = 0
        row += 1
workbook.close() 

workbook = xlsxwriter.Workbook('Comments.xlsx') 
worksheet = workbook.add_worksheet() 
row = 0
column = 0

for item in ArrayCommnent: 
    worksheet.write(row, column, item)
    column += 1
    if (column  % 3) == 0:
        column = 0
        row += 1
workbook.close()

workbook = xlsxwriter.Workbook('Flairs.xlsx') 
worksheet = workbook.add_worksheet() 
row = 0
column = 0

for item in ArrayFlair: 
    worksheet.write(row, column, item)
    column += 1
    if (column  % 2) == 0:
        column = 0
        row += 1
workbook.close()

try:
    print('Total Post Upvotes: '+str(TotalPostUpvotes))
    print('Total Post Downvotes: '+str(TotalPostDownvotes))
    print('Total Post Average Upvotes: '+str((round(TotalPostUpvotes/PostQuantity,2))))
    print('Total Post Average Downvotes: '+str((round(TotalPostDownvotes/PostQuantity,2))))
    print('Comment Quantity: '+str(CommentQuantity))
    print('Post Quantity: '+str(PostQuantity))
    print('Total Comment Upvotes: '+str(TotalCommentUpvotes))
    print('Average Upvotes Comment: '+str((round(TotalCommentUpvotes/PostQuantity,2))))
    print()
    print('The worst comment is:')
    print(WorstComment.translate(non_bmp_map))
    print()
    print('With '+str(WorstCommentUpvote)+' upvotes of the author: '+WorstAuthor+' of the post: '+WorstCommentPost.translate(non_bmp_map))
    print()
    print('The best comment is: ')
    print(BestComment.translate(non_bmp_map))
    print()
    print('With '+str(BestCommentUpvote)+' upvotes of the author: '+BestAuthor+' of the post: '+BestCommentPost.translate(non_bmp_map))
    print()
    FinalTime = time() 
    ExecutionTime = FinalTime  - InitialTime
    print ('Execution Time:',round(ExecutionTime ,2),'seconds ó',round(ExecutionTime/60,2),'minutes') 
except UnicodeEncodeError:
    print('Unicode Encode Error')
