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

def GenerateDocuments():
    GeneratePostDocument()
    GenerateCommentDocument()
    GenerateFlairDocument()
    GenerateStatisticDocument()
    
def GeneratePostDocument():
    workbook = xlsxwriter.Workbook('Posts.xlsx') 
    worksheet = workbook.add_worksheet() 
    row = 0
    column = 0

    worksheet.write(row, column, "User")
    worksheet.write(row, column+1, "Upvotes")
    worksheet.write(row, column+2, "Downvotes Posts")
    worksheet.write(row, column+3, "Quantity Posts")

    for item in ArrayPost: 
        worksheet.write(row+1, column, item)
        column += 1
        if (column% 4) == 0:
            column = 0
            row += 1
    workbook.close()
    
def GenerateCommentDocument():
    workbook = xlsxwriter.Workbook('Comments.xlsx') 
    worksheet = workbook.add_worksheet()
    row = 0
    column = 0

    worksheet.write(row, column, "User")
    worksheet.write(row, column+1, "Upvotes")
    worksheet.write(row, column+2, "Quantity Comments")

    for item in ArrayCommnent: 
        worksheet.write(row+1, column, item)
        column += 1
        if (column  % 3) == 0:
            column = 0
            row += 1
    workbook.close()
    
def GenerateFlairDocument():
    workbook = xlsxwriter.Workbook('Flairs.xlsx') 
    worksheet = workbook.add_worksheet() 
    row = 0
    column = 0

    worksheet.write(row, column, "Flairs")
    worksheet.write(row, column+1, "Quantity")

    for item in ArrayFlair: 
        worksheet.write(row+1, column, item)
        column += 1
        if (column  % 2) == 0:
            column = 0
            row += 1
    workbook.close()
    
def GenerateStatisticDocument():
    workbook = xlsxwriter.Workbook('Statistics.xlsx') 
    worksheet = workbook.add_worksheet()
    row = 0
    column = 0

    worksheet.write(row, column, "Statistics")
    worksheet.write(row, column+1, "Quantity")
    worksheet.write(row+1, column, "Total Post Upvotes")
    worksheet.write(row+1, column+1, TotalPostUpvotes)
    worksheet.write(row+2, column, "Total Post Downvotes")
    worksheet.write(row+2, column+1, TotalPostDownvotes)
    worksheet.write(row+3, column, "Total Post Average Upvotes")
    worksheet.write(row+3, column+1, round(TotalPostUpvotes/PostQuantity,2))
    worksheet.write(row+4, column, "Total Post Average Downvote")
    worksheet.write(row+4, column+1, round(TotalPostDownvotes/PostQuantity,2))
    worksheet.write(row+5, column, "Comment Quantity")
    worksheet.write(row+5, column+1, CommentQuantity)
    worksheet.write(row+6, column, "Post Quantity")
    worksheet.write(row+6, column+1, PostQuantity)
    worksheet.write(row+7, column, "Total Comment Upvotes")
    worksheet.write(row+7, column+1, TotalCommentUpvotes)
    worksheet.write(row+8, column, "Average Upvotes Comment")
    worksheet.write(row+8, column+1, round(TotalCommentUpvotes/PostQuantity,2))
    worksheet.write(row+9, column, "The worst comment is")
    worksheet.write(row+9, column+1, WorstComment.translate(non_bmp_map))
    worksheet.write(row+10, column, "With")
    worksheet.write(row+10, column+1, WorstCommentUpvote)
    worksheet.write(row+11, column, "upvotes of the author")
    worksheet.write(row+11, column+1, WorstAuthor)
    worksheet.write(row+12, column, "of the post")
    worksheet.write(row+12, column+1, WorstCommentPost.translate(non_bmp_map))
    worksheet.write(row+13, column, "The best comment is")
    worksheet.write(row+13, column+1, BestComment.translate(non_bmp_map))
    worksheet.write(row+14, column, "With")
    worksheet.write(row+14, column+1, BestCommentUpvote)
    worksheet.write(row+15, column, "upvotes of the author")
    worksheet.write(row+15, column+1, BestAuthor)
    worksheet.write(row+16, column, "of the post")
    worksheet.write(row+16, column+1, BestCommentPost.translate(non_bmp_map))
    workbook.close()
    
GenerateDocuments()
FinalTime = time() 
ExecutionTime = FinalTime  - InitialTime
print ('Execution Time:',round(ExecutionTime ,2),'seconds ó',round(ExecutionTime/60,2),'minutes') 