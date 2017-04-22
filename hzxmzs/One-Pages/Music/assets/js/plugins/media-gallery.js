function mySwiper1() {
  var swiper1 = new Swiper('.swiper-images', {
    pagination: '.swiper-pagination-images',
    slidesPerView: 6,
    slidesPerColumn: 2,
    spaceBetween: 30,
    paginationClickable: true,
    breakpoints: {
      // when window width is <= 991px
      2000: {
        slidesPerView: 6,
      },
      // when window width is <= 991px
      991: {
        slidesPerView: 5,
      },
      // when window width is <= 767px
      767: {
        slidesPerView: 4,
      },
      // when window width is <= 599px
      599: {
        slidesPerView: 3,
      },
      // when window width is <= 499px
      499: {
        slidesPerView: 2,
      },
      // when window width is <= 399px
      399: {
        slidesPerView: 1,
      },
    }
  });
}
function mySwiper2() {
  var swiper2 = new Swiper('.swiper-videos', {
    pagination: '.swiper-pagination-videos',
    slidesPerView: 6,
    slidesPerColumn: 2,
    spaceBetween: 30,
    paginationClickable: true,
    breakpoints: {
      // when window width is <= 991px
      2000: {
        slidesPerView: 6,
      },
      // when window width is <= 991px
      991: {
        slidesPerView: 5,
      },
      // when window width is <= 767px
      767: {
        slidesPerView: 4,
      },
      // when window width is <= 599px
      599: {
        slidesPerView: 3,
      },
      // when window width is <= 499px
      499: {
        slidesPerView: 2,
      },
      // when window width is <= 399px
      399: {
        slidesPerView: 1,
      },
    }
  });
}
mySwiper1();
$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
  e.target;
  e.relatedTarget;
  mySwiper2();
})