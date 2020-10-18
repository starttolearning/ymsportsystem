from ymsports_summary import ymsports_summary
from ymsports_summary_rank import ymsports_summary_rank


if __name__ == '__main__':
    # define input and outp file
    input_filename = 'YMS单项详细记录表7-9'
    class_counts = [13, 15, 14]

    your_decision = input("Do you want have a RANKED list? (Y/n)")
    if your_decision.lower() == 'y':
        output_filename = f'全校排名RANK1-4'
        ys = ymsports_summary_rank(
            input_filename, output_filename, class_counts)
    else:
        output_filename = f'YMS总分表7-8'
        ys = ymsports_summary(input_filename, output_filename, class_counts)
